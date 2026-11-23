# -*- coding: utf-8 -*-
"""
멀티시트 엑셀(.xlsx) 생성 및 서식 적용 모듈 (openpyxl 기반)

시트 구성:
  1) [전체_통합]
  2) [충청남도]
  3) [충청북도]
  4) [대전광역시]
  5) [전북특별자치도]

적용 서식:
  - 1행 헤더 배경색 지정 및 틀 고정 (Freeze Panes: A2)
  - 전체 컬럼 자동 필터 (AutoFilter)
  - 한글(동아시아 광폭 문자) 글자 폭을 반영한 열 너비 자동 조절
"""

from __future__ import annotations

import io
import unicodedata
from typing import Dict, List, Optional, Tuple

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from processor import (
    CATEGORY_BOTH,
    CATEGORY_SEASONED,
    CATEGORY_SLICED,
    OUTPUT_COLUMNS,
)

# 시트명 및 시/도 필터 매핑 (총 5개 시트 고정)
SHEET_DEFINITIONS: Tuple[Tuple[str, Optional[str]], ...] = (
    ("전체_통합", None),
    ("충청남도", "충청남도"),
    ("충청북도", "충청북도"),
    ("대전광역시", "대전광역시"),
    ("전북특별자치도", "전북특별자치도(전라북도)"),
)

# 열별 최소·최대 너비 가이드 (초장문 생산품목 리스트 등으로 열이 과도하게 넓어지는 현상 방지)
COLUMN_WIDTH_BOUNDS: Dict[str, Tuple[int, int]] = {
    "시/도": (14, 26),
    "시/군/구": (12, 20),
    "가공구분": (22, 30),
    "업체명": (18, 38),
    "대표자": (10, 18),
    "인허가업종": (18, 36),
    "연락처(전화번호)": (15, 22),
    "소재지 주소": (28, 56),
    "주요 품목유형": (16, 36),
    "대표 생산품목 리스트": (24, 65),
    "등록 품목수": (12, 16),
    "인허가번호": (15, 28),
    "허가일자": (13, 16),
}


def _visual_text_length(value: object) -> float:
    """한글·전각 문자는 1.9배, 영문·숫자는 1.0배로 시각적 문자열 폭을 계산한다."""
    if value is None:
        return 0.0
    text = str(value)
    if not text:
        return 0.0
    # 줄바꿈이 있을 경우 가장 긴 줄 기준
    max_line_width = 0.0
    for line in text.splitlines():
        width = 0.0
        for ch in line:
            eaw = unicodedata.east_asian_width(ch)
            width += 1.9 if eaw in ("W", "F") else 1.0
        if width > max_line_width:
            max_line_width = width
    return max_line_width


def _write_styled_sheet(ws: Worksheet, df: pd.DataFrame) -> None:
    """단일 워크시트에 헤더·데이터 기록 및 서식(배경색, 틀 고정, 자동필터, 열 너비)을 적용한다."""
    columns: List[str] = list(OUTPUT_COLUMNS)
    ordered_df = df.reindex(columns=columns) if not df.empty else pd.DataFrame(columns=columns)

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="맑은 고딕", size=11, bold=True, color="FFFFFF")
    body_font = Font(name="맑은 고딕", size=10, color="111111")

    cat_fills = {
        CATEGORY_BOTH: PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"),
        CATEGORY_SEASONED: PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid"),
        CATEGORY_SLICED: PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"),
    }

    thin_side = Side(border_style="thin", color="D9D9D9")
    header_border = Border(
        left=Side(border_style="thin", color="FFFFFF"),
        right=Side(border_style="thin", color="FFFFFF"),
        top=Side(border_style="thin", color="1F4E78"),
        bottom=Side(border_style="medium", color="0F243E"),
    )
    cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    align_header = Alignment(horizontal="center", vertical="center", wrap_text=False)
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=False)
    align_right = Alignment(horizontal="right", vertical="center", wrap_text=False)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=False)
    align_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)

    center_cols = {"시/도", "시/군/구", "가공구분", "대표자", "연락처(전화번호)", "인허가번호", "허가일자"}
    wrap_cols = {"대표 생산품목 리스트", "주요 품목유형"}

    # 열 너비 계산용 최대 길이 추적 (헤더는 자동필터 드롭다운 화살표 여유분 +4 반영)
    max_widths: List[float] = [_visual_text_length(col) + 4.5 for col in columns]

    # 1행 헤더 작성
    ws.row_dimensions[1].height = 26
    for col_idx, col_name in enumerate(columns, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = align_header
        cell.border = header_border

    # 데이터 행 작성
    for row_idx, row_tuple in enumerate(ordered_df.itertuples(index=False, name=None), start=2):
        ws.row_dimensions[row_idx].height = 21
        for col_idx, (col_name, val) in enumerate(zip(columns, row_tuple), start=1):
            if pd.isna(val):
                clean_val: object = ""
            elif col_name == "등록 품목수":
                clean_val = int(val)
            else:
                clean_val = str(val)

            cell = ws.cell(row=row_idx, column=col_idx, value=clean_val)
            cell.font = body_font
            cell.border = cell_border

            if col_name == "등록 품목수":
                cell.alignment = align_right
                cell.number_format = "#,##0"
            elif col_name in center_cols:
                cell.alignment = align_center
            elif col_name in wrap_cols:
                cell.alignment = align_wrap
            else:
                cell.alignment = align_left

            if col_name == "가공구분" and isinstance(clean_val, str) and clean_val in cat_fills:
                cell.fill = cat_fills[clean_val]

            v_len = _visual_text_length(clean_val) + 2.5
            if v_len > max_widths[col_idx - 1]:
                max_widths[col_idx - 1] = v_len

    # 1행 틀 고정 (Freeze Panes)
    ws.freeze_panes = "A2"

    # 전체 컬럼 자동 필터 적용
    last_col_letter = get_column_letter(len(columns))
    last_row_num = max(1, len(ordered_df) + 1)
    ws.auto_filter.ref = f"A1:{last_col_letter}{last_row_num}"

    # 한글 길이를 반영한 열 너비 자동 조절
    for col_idx, col_name in enumerate(columns, start=1):
        col_letter = get_column_letter(col_idx)
        min_w, max_w = COLUMN_WIDTH_BOUNDS.get(col_name, (12, 50))
        computed_width = max(min_w, min(max_w, round(max_widths[col_idx - 1], 1)))
        ws.column_dimensions[col_letter].width = computed_width


def generate_multisheet_excel_bytes(df: pd.DataFrame) -> bytes:
    """
    정제된 DataFrame을 받아 5개 시트([전체_통합], [충청남도], [충청북도], [대전광역시], [전북특별자치도])가
    포함된 서식 적용 엑셀(.xlsx) 바이너리를 반환한다.
    """
    wb = Workbook()
    # 기본 생성 시트 제거 후 정의된 5개 시트 순서대로 생성
    default_ws = wb.active
    if default_ws is not None:
        wb.remove(default_ws)

    safe_df = df.copy() if df is not None and not df.empty else pd.DataFrame(columns=OUTPUT_COLUMNS)

    for sheet_title, sido_filter in SHEET_DEFINITIONS:
        ws = wb.create_sheet(title=sheet_title)
        if sido_filter is None:
            sub_df = safe_df
        else:
            sub_df = safe_df[safe_df["시/도"] == sido_filter].reset_index(drop=True)
        _write_styled_sheet(ws, sub_df)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()
