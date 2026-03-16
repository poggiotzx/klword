# -*- coding: utf-8 -*-
"""klword 对外导出。"""

from .word_api import DocContainer, WordAPI
from .word_styles import (
    BODY_STYLE,
    CAPTION_STYLE,
    MAIN_STYLE,
    TABLE_BODY_STYLE,
    TABLE_HEADER_STYLE,
    CellStyle,
    RichTextValue,
    TableStyle,
    TextStyle,
    make_rich_text,
    make_table_style,
)

__all__ = [
    "WordAPI",
    "DocContainer",
    "TextStyle",
    "CellStyle",
    "TableStyle",
    "RichTextValue",
    "BODY_STYLE",
    "CAPTION_STYLE",
    "MAIN_STYLE",
    "TABLE_BODY_STYLE",
    "TABLE_HEADER_STYLE",
    "make_rich_text",
    "make_table_style",
]
