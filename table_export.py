"""표 → 엑셀(xlsx) 내보내기 — 검은 헤더 스타일, 화면과 같은 서식.

발송성과·주간보고 두 대시보드가 같이 쓴다. Streamlit 기본 툴바의 'Download as CSV'는
서식이 다 날아가고 한글 인코딩도 흔들려서, 보고서에 바로 붙일 수 있는 표를 따로 만든다.

핵심은 **값을 문자열이 아니라 '숫자 + 엑셀 표시형식'으로 넣는 것** — 그래야 받는 쪽에서
정렬·합계가 된다. 화면 표시 문자열(천단위·%·부호)에서 표시형식을 역으로 추론한다.
"""
import io
import re

import numpy as np
import pandas as pd


XL_HEAD_BG = "111827"      # 거의 검정 (헤더)
XL_HEAD_FG = "FFFFFF"
XL_ZEBRA = "F6F7F9"
XL_BORDER = "D9DDE3"
_XL_PCT_RE = re.compile(r"^([+-]?)[\d,]*\.?(\d*)\s*%$")
_XL_NUM_RE = re.compile(r"^([+-]?)[\d,]+\.?(\d*)$")


def _xl_numfmt(vals):
    """화면 표시 문자열들 → 엑셀 숫자 서식. 숫자로 볼 수 없으면 None(=문자로 씀)."""
    vals = [str(v).strip() for v in vals if str(v).strip() not in ("", "–", "-", "nan", "None")]
    if not vals:
        return None
    pcts = [_XL_PCT_RE.match(v) for v in vals]
    if all(pcts):
        dec = max(len(m.group(2)) for m in pcts)
        base = "0." + "0" * dec + "%" if dec else "0%"
        return f"+{base};-{base}" if any(m.group(1) == "+" for m in pcts) else base
    nums = [_XL_NUM_RE.match(v) for v in vals]
    if all(nums):
        dec = max(len(m.group(2)) for m in nums)
        base = ("#,##0." + "0" * dec) if dec else "#,##0"
        return f"+{base};-{base}" if any(m.group(1) == "+" for m in nums) else base
    return None


def _xl_display(obj):
    """Styler든 DataFrame이든 (원본 df, 화면 표시 문자열 df)로 돌려준다."""
    df = getattr(obj, "data", obj)
    if not isinstance(df, pd.DataFrame):
        df = pd.DataFrame(df)
    disp = df.astype(object).map(lambda v: "" if pd.isna(v) else str(v))
    if hasattr(obj, "_translate"):                        # Styler — 화면과 같은 서식으로
        try:
            body = obj._translate(False, False)["body"]
            rows = [[c.get("display_value") for c in r if c.get("type") == "td"] for r in body]
            rows = [r for r in rows if r]
            if len(rows) == len(df) and all(len(r) == df.shape[1] for r in rows):
                disp = pd.DataFrame(rows, index=df.index, columns=df.columns)
        except Exception:                                 # noqa: BLE001
            pass                                          # 실패해도 원본 문자열로 내보낸다
    return df, disp


def xlsx_bytes(obj, sheet_name="데이터", title=None, index=False):
    """표 → 검은 헤더 스타일 xlsx 바이트. Styler·DataFrame 모두 받는다."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    df, disp = _xl_display(obj)
    if index:
        df = df.reset_index()
        disp = disp.reset_index()
    cols = ["" if c is None else str(c) for c in df.columns]

    wb = Workbook()
    ws = wb.active
    ws.title = re.sub(r"[\[\]:*?/\\]", " ", str(sheet_name))[:31] or "데이터"
    r0 = 1
    if title:
        ws.cell(row=1, column=1, value=str(title)).font = Font(bold=True, size=13, color="111827")
        r0 = 3
    thin = Side(style="thin", color=XL_BORDER)
    bd = Border(left=thin, right=thin, top=thin, bottom=thin)
    for j, c in enumerate(cols, start=1):
        cell = ws.cell(row=r0, column=j, value=c)
        cell.fill = PatternFill("solid", fgColor=XL_HEAD_BG)
        cell.font = Font(bold=True, color=XL_HEAD_FG, size=11)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = bd
    fmts = [_xl_numfmt(disp.iloc[:, j].tolist()) for j in range(df.shape[1])]
    zebra = PatternFill("solid", fgColor=XL_ZEBRA)
    for i in range(len(df)):
        for j in range(df.shape[1]):
            raw, txt = df.iat[i, j], disp.iat[i, j]
            cell = ws.cell(row=r0 + 1 + i, column=j + 1)
            if fmts[j] is not None and isinstance(raw, (int, float, np.integer, np.floating)) \
                    and not pd.isna(raw):
                cell.value = float(raw)
                cell.number_format = fmts[j]
                cell.alignment = Alignment(horizontal="right")
            else:
                cell.value = "" if txt in ("nan", "None") else txt
                cell.alignment = Alignment(horizontal="left", vertical="center")
            cell.font = Font(size=10, color="1F2937")
            cell.border = bd
            if i % 2 == 1:
                cell.fill = zebra
    for j, c in enumerate(cols, start=1):
        width = max([len(str(c))] + [len(str(x)) for x in disp.iloc[:, j - 1].tolist()[:200]])
        ws.column_dimensions[get_column_letter(j)].width = min(max(width * 1.6 + 2, 9), 42)
    ws.row_dimensions[r0].height = 24
    ws.freeze_panes = ws.cell(row=r0 + 1, column=1)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
