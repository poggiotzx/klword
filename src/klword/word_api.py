# -*- coding: utf-8 -*-
"""
核心 Word API

Copyright (c) 2026
Released under the MIT License.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Optional

from docx import Document
from docx.document import Document as _DocumentType
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from .word_styles import (
    BODY_STYLE,
    CAPTION_STYLE,
    TABLE_BODY_STYLE,
    TABLE_HEADER_STYLE,
    CellStyle,
    RichTextValue,
    TableStyle,
    TextStyle,
    make_rich_text,
)

IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp", ".tif", ".tiff"}


@dataclass(slots=True)
class ContainerOp:
    kind: str
    payload: dict[str, Any]


class WordAPI:
    """Word 文档生成入口。"""

    def __init__(self, template_path: str) -> None:
        template = Path(template_path)
        if not template.exists():
            raise FileNotFoundError(f"模板不存在: {template}")
        self.template_path = str(template)
        self._table_index = 0
        self._figure_index = 0

    def new_container(self) -> "DocContainer":
        """创建一个复合内容容器。"""
        return DocContainer(self)

    def next_table_caption(self) -> int:
        self._table_index += 1
        return self._table_index

    def next_figure_caption(self) -> int:
        self._figure_index += 1
        return self._figure_index

    def render(self, context: dict[str, Any], output_path: str) -> str:
        """基于模板与上下文渲染输出文档。"""
        doc = Document(self.template_path)
        paragraphs = list(doc.paragraphs)
        for paragraph in paragraphs:
            raw_text = paragraph.text.strip()
            if not raw_text:
                continue

            key = self._extract_key(raw_text)
            if not key or key not in context:
                continue

            value = context[key]
            if isinstance(value, RichTextValue):
                self._replace_text_placeholder(paragraph, value)
            elif isinstance(value, DocContainer):
                self._replace_container_placeholder(doc, paragraph, value)
            else:
                self._replace_text_placeholder(paragraph, make_rich_text(str(value), BODY_STYLE))

        out_path = Path(output_path)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(out_path))
        return str(out_path)

    def write_header_footer(self, docx_path: str, header_text: str = "") -> None:
        """写入简单页眉。"""
        doc = Document(docx_path)
        for section in doc.sections:
            header = section.header
            paragraph = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.text = ""
            run = paragraph.add_run(header_text)
            run.font.name = "宋体"
            run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "宋体")
            run.font.size = Pt(10.5)
        doc.save(docx_path)

    @staticmethod
    def _extract_key(text: str) -> Optional[str]:
        if text.startswith("{{p ") and text.endswith(" }}"):
            return text[4:-3].strip()
        if text.startswith("{{ ") and text.endswith(" }}"):
            return text[3:-3].strip()
        if text.startswith("{{") and text.endswith("}}"):
            return text[2:-2].strip()
        return None

    @staticmethod
    def _clear_paragraph(paragraph) -> None:
        p = paragraph._p
        for child in list(p):
            p.remove(child)

    def _replace_text_placeholder(self, paragraph, value: RichTextValue) -> None:
        self._clear_paragraph(paragraph)
        run = paragraph.add_run(value.text)
        _apply_text_style(paragraph, run, value.style)

    def _replace_container_placeholder(self, doc: _DocumentType, paragraph, container: "DocContainer") -> None:
        anchor = paragraph._p
        parent = anchor.getparent()
        current_anchor = anchor
        for op in container.ops:
            current_anchor = _insert_op_after(doc, current_anchor, op)
        parent.remove(anchor)


class DocContainer:
    """复合内容容器。"""

    def __init__(self, api: WordAPI) -> None:
        self.api = api
        self.ops: list[ContainerOp] = []
        self.subdoc = self

    def add_title(self, text: str, style: TextStyle) -> None:
        self.ops.append(ContainerOp("paragraph", {"text": text, "style": style}))

    def add_paragraph(self, text: str, style: TextStyle) -> None:
        self.ops.append(ContainerOp("paragraph", {"text": text, "style": style}))

    def add_image(self, image_path: str, width_cm: float = 8.0, align: str = "center") -> None:
        image = Path(image_path)
        if not image.exists():
            raise FileNotFoundError(f"图片不存在: {image}")
        self.ops.append(ContainerOp("image", {"image_path": str(image), "width_cm": width_cm, "align": align}))

    def add_table_caption_auto(self, title: str, style: TextStyle = CAPTION_STYLE) -> None:
        index = self.api.next_table_caption()
        self.add_paragraph(f"表 {index}  {title}", style)

    def add_figure_caption_auto(self, title: str, style: TextStyle = CAPTION_STYLE) -> None:
        index = self.api.next_figure_caption()
        self.add_paragraph(f"图 {index}  {title}", style)

    def add_table_by_config(self, config: dict[str, Any]) -> None:
        self.ops.append(ContainerOp("table", {"config": config}))



def _insert_op_after(doc: _DocumentType, anchor, op: ContainerOp):
    if op.kind == "paragraph":
        payload = op.payload
        p = doc.add_paragraph()
        run = p.add_run(payload["text"])
        _apply_text_style(p, run, payload["style"])
        return _move_element_after(p._p, anchor)

    if op.kind == "image":
        payload = op.payload
        p = doc.add_paragraph()
        p.alignment = _to_alignment(payload["align"])
        run = p.add_run()
        run.add_picture(str(payload["image_path"]), width=Cm(payload["width_cm"]))
        return _move_element_after(p._p, anchor)

    if op.kind == "table":
        tbl = _build_table(doc, op.payload["config"])
        return _move_element_after(tbl._tbl, anchor)

    raise ValueError(f"Unsupported op kind: {op.kind}")



def _move_element_after(element, anchor):
    parent = anchor.getparent()
    parent.remove(element)
    anchor.addnext(element)
    return element



def _build_table(doc: _DocumentType, config: dict[str, Any]):
    data = config.get("data", [])
    if not data:
        table = doc.add_table(rows=1, cols=1)
        table.cell(0, 0).text = ""
        return table

    style_cfg = config.get("style", {})
    header_style = style_cfg.get("header", TABLE_HEADER_STYLE)
    body_style = style_cfg.get("body", TABLE_BODY_STYLE)
    table_style = style_cfg.get("table", TableStyle())
    if not isinstance(table_style, TableStyle):
        table_style = TableStyle(
            col_widths_cm=table_style.get("col_widths_cm"),
            row_heights_cm=table_style.get("row_heights_cm"),
            image_width_cm=table_style.get("image_width_cm", 8.0),
        )

    rows = len(data)
    cols = max(len(row) for row in data)
    table = doc.add_table(rows=rows, cols=cols)
    try:
        table.style = "Table Grid"
    except KeyError:
        pass
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    if table_style.col_widths_cm:
        for col_idx, width_cm in enumerate(table_style.col_widths_cm):
            if col_idx >= cols:
                break
            for cell in table.columns[col_idx].cells:
                cell.width = Cm(width_cm)

    for row_idx, row_data in enumerate(data):
        row = table.rows[row_idx]
        for col_idx in range(cols):
            cell = row.cells[col_idx]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            value = row_data[col_idx] if col_idx < len(row_data) else ""
            _fill_cell(
                cell=cell,
                value=value,
                style=header_style if row_idx == 0 else body_style,
                image_width_cm=table_style.image_width_cm,
            )
    return table



def _fill_cell(cell, value: Any, style: CellStyle, image_width_cm: float) -> None:
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.alignment = _to_alignment(style.align)
    if _is_image_path(value):
        run = paragraph.add_run()
        run.add_picture(str(value), width=Cm(image_width_cm))
    else:
        run = paragraph.add_run(str(value))
        _apply_text_style(paragraph, run, style, allow_indent=False)



def _is_image_path(value: Any) -> bool:
    if not isinstance(value, (str, Path)):
        return False
    return Path(str(value)).suffix.lower() in IMAGE_EXTS



def _to_alignment(value: str) -> int:
    mapping = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }
    return mapping.get(value, WD_ALIGN_PARAGRAPH.LEFT)



def _apply_text_style(paragraph, run, style: TextStyle | CellStyle, allow_indent: bool = True) -> None:
    paragraph.alignment = _to_alignment(style.align)
    if style.space_before_pt:
        paragraph.paragraph_format.space_before = Pt(style.space_before_pt)
    if style.space_after_pt:
        paragraph.paragraph_format.space_after = Pt(style.space_after_pt)
    if style.line_spacing:
        paragraph.paragraph_format.line_spacing = style.line_spacing
    if allow_indent and getattr(style, "first_line_indent_chars", 0.0):
        paragraph.paragraph_format.first_line_indent = Pt(style.font_size_pt * style.first_line_indent_chars)

    run.font.name = style.font_name
    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), style.font_name)
    run.font.size = Pt(style.font_size_pt)
    run.bold = style.bold
    run.italic = style.italic
    run.underline = style.underline
    if style.color_rgb:
        run.font.color.rgb = RGBColor(*style.color_rgb)
