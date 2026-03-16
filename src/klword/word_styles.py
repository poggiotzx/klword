# -*- coding: utf-8 -*-
"""
样式定义与样式工厂

Copyright (c) 2026
Released under the MIT License.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Optional, Sequence

FONT_SIZE_MAP = {
    "初号": 42.0,
    "小初": 36.0,
    "一号": 26.0,
    "小一": 24.0,
    "二号": 22.0,
    "小二": 18.0,
    "三号": 16.0,
    "小三": 15.0,
    "四号": 14.0,
    "小四": 12.0,
    "五号": 10.5,
    "小五": 9.0,
    "六号": 7.5,
}


def _to_pt(value: str | float | int | None, default: float = 12.0) -> float:
    if value is None:
        return default
    if isinstance(value, (int, float)):
        return float(value)
    return FONT_SIZE_MAP.get(value, default)


@dataclass(slots=True)
class TextStyle:
    font_name: str = "宋体"
    font_size: str | float = "小四"
    bold: bool = False
    italic: bool = False
    underline: bool = False
    color_rgb: Optional[tuple[int, int, int]] = None
    align: str = "left"
    first_line_indent_chars: float = 0.0
    space_before_pt: float = 0.0
    space_after_pt: float = 0.0
    line_spacing: float = 1.5

    @property
    def font_size_pt(self) -> float:
        return _to_pt(self.font_size)


@dataclass(slots=True)
class CellStyle(TextStyle):
    vertical_align: str = "center"


@dataclass(slots=True)
class TableStyle:
    col_widths_cm: Optional[Sequence[float]] = None
    row_heights_cm: Optional[Sequence[float]] = None
    image_width_cm: float = 8.0


@dataclass(slots=True)
class RichTextValue:
    text: str
    style: TextStyle


MAIN_STYLE = TextStyle(font_name="黑体", font_size="二号", bold=True, align="center", space_after_pt=12)
BODY_STYLE = TextStyle(font_name="宋体", font_size="小四", line_spacing=1.5, first_line_indent_chars=2)
CAPTION_STYLE = TextStyle(font_name="宋体", font_size="五号", align="center", space_before_pt=6, space_after_pt=6)
TABLE_HEADER_STYLE = CellStyle(font_name="宋体", font_size="小四", bold=True, align="center")
TABLE_BODY_STYLE = CellStyle(font_name="宋体", font_size="小四", align="center")


def make_rich_text(text: str, style: TextStyle) -> RichTextValue:
    """创建富文本值。"""
    return RichTextValue(text=text, style=style)



def make_table_style(
    col_widths_cm: Optional[Sequence[float]] = None,
    row_heights_cm: Optional[Sequence[float]] = None,
    image_width_cm: float = 8.0,
) -> TableStyle:
    """创建表格样式配置。"""
    return TableStyle(
        col_widths_cm=col_widths_cm,
        row_heights_cm=row_heights_cm,
        image_width_cm=image_width_cm,
    )
