"""見積依頼書(Excel)の生成。元フォーマット「リフォーム・リノベーション工事見積書」の体裁を踏襲する。

単価は業者が入力する前提のため空欄にし、金額・小計・合計は数式にしてある。
"""
from io import BytesIO

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

FONT = "ＭＳ 明朝"
THIN = Side(style="thin")
HAIR = Side(style="hair")
HEADER_FILL = PatternFill("solid", fgColor="D9D9D9")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")  # 業者の入力欄

# A:No B:名称 C:仕様 D:数量 E:(余白) F:単位 G:単価 H:(余白) I:金額 J:(余白) K:摘要
COL_WIDTHS = {"A": 6, "B": 38, "C": 34, "D": 9, "E": 2, "F": 7, "G": 14, "H": 2, "I": 16, "J": 2, "K": 28}
LAST_COL = 11


def _f(size=11, bold=False):
    return Font(name=FONT, size=size, bold=bold)


def _qty(v):
    if v is None or v == "":
        return None
    return int(v) if float(v) == int(v) else round(float(v), 2)


def build_workbook(data: dict) -> BytesIO:
    prop = data["property"]
    cats = data["categories"]
    wb = Workbook()
    ws = wb.active
    ws.title = "見積書"
    for col, w in COL_WIDTHS.items():
        ws.column_dimensions[col].width = w

    def box(row, c1=1, c2=LAST_COL):
        for c in range(c1, c2 + 1):
            ws.cell(row, c).border = Border(left=THIN if c in (c1, 2, 3, 4, 6, 7, 9, 11) else None,
                                            right=THIN if c in (c2, 1, 2, 3, 5, 6, 8, 10) else None,
                                            top=HAIR, bottom=HAIR)

    # ---- タイトル ----
    title = f"{prop.get('customer_name') or 'お名前'} 様邸{prop.get('work_title') or 'リノベーション工事'} 見積書"
    ws.cell(1, 2, title).font = _f(14, True)
    ws.row_dimensions[1].height = 24
    ws.merge_cells(start_row=2, start_column=7, end_row=2, end_column=11)
    ws.cell(2, 2, f"{prop['vendor_name']} 御中" if prop.get("vendor_name") else "").font = _f()
    ws.cell(2, 7, prop.get("company_name") or "").font = _f()
    ws.cell(2, 7).alignment = Alignment(horizontal="right")

    # ---- ヘッダ ----
    hr = 3
    heads = {1: "No.", 2: "名　　　称", 3: "仕　　　様", 4: "数量", 6: "単位", 7: "単　価", 9: "金　額", 11: "摘　要"}
    for c in range(1, LAST_COL + 1):
        cell = ws.cell(hr, c)
        cell.fill = HEADER_FILL
        cell.border = Border(top=THIN, bottom=THIN, left=THIN if c in (1, 2, 3, 4, 6, 7, 9, 11) else None,
                             right=THIN if c == LAST_COL else None)
    for c, t in heads.items():
        ws.cell(hr, c, t).font = _f(bold=True)
        ws.cell(hr, c).alignment = Alignment(horizontal="center", vertical="center")
    for a, b in ((4, 5), (7, 8), (9, 10)):
        ws.merge_cells(start_row=hr, start_column=a, end_row=hr, end_column=b)

    row = hr + 1
    ws.cell(row, 1, "（内　訳　明　細）").font = _f()
    row += 1

    # 大項目サマリ(金額は後で小計セルを参照させる)
    summary_rows = []
    for i, cat in enumerate(cats, 1):
        ws.cell(row, 1, i).font = _f()
        ws.cell(row, 2, cat["name"]).font = _f()
        ws.cell(row, 4, 1).font = _f()
        ws.cell(row, 6, "式").font = _f()
        box(row)
        summary_rows.append(row)
        row += 1
    adj_row = row + 1
    ws.cell(adj_row, 2, "調整費").font = _f()
    ws.cell(adj_row, 9).fill = INPUT_FILL
    ws.cell(adj_row, 9).number_format = '#,##0;[Red]-#,##0'
    ws.cell(adj_row, 9).font = _f()
    box(adj_row)
    total_row = adj_row + 2
    ws.cell(total_row, 2, "合　計").font = _f(bold=True)
    for c in range(1, LAST_COL + 1):
        ws.cell(total_row, c).border = Border(top=THIN, bottom=THIN, left=THIN if c in (1, 2, 3, 4, 6, 7, 9, 11) else None,
                                             right=THIN if c == LAST_COL else None)
    row = total_row + 2

    # ---- 各大項目の内訳 ----
    subtotal_cells = []
    for i, cat in enumerate(cats, 1):
        ws.cell(row, 1, "（内　訳　明　細）").font = _f()
        row += 1
        ws.cell(row, 1, i).font = _f(bold=True)
        ws.cell(row, 2, cat["name"]).font = _f(bold=True)
        box(row)
        row += 1
        first = row
        for it in cat["items"]:
            ws.cell(row, 2, it["name"]).font = _f(bold=bool(it.get("group")))
            if not it.get("group"):
                ws.cell(row, 3, it.get("spec") or None).font = _f()
                q = ws.cell(row, 4, _qty(it.get("qty")))
                q.font = _f()
                q.alignment = Alignment(horizontal="right")
                q.number_format = "#,##0.##"
                ws.cell(row, 6, it.get("unit") or None).font = _f()
                ws.cell(row, 6).alignment = Alignment(horizontal="center")
                g = ws.cell(row, 7)
                g.fill = INPUT_FILL
                g.number_format = "#,##0"
                g.font = _f()
                a = ws.cell(row, 9, f'=IF(G{row}="","",ROUND(D{row}*G{row},0))')
                a.number_format = "#,##0"
                a.font = _f()
                ws.cell(row, 11, it.get("note") or None).font = _f()
                ws.cell(row, 11).fill = INPUT_FILL
            box(row)
            row += 1
        last = row - 1
        ws.cell(row, 2, "小　計").font = _f(bold=True)
        s = ws.cell(row, 9, f"=SUM(I{first}:I{last})")
        s.number_format = "#,##0"
        s.font = _f(bold=True)
        for c in range(1, LAST_COL + 1):
            ws.cell(row, c).border = Border(top=THIN, bottom=THIN, left=THIN if c in (1, 2, 3, 4, 6, 7, 9, 11) else None,
                                            right=THIN if c == LAST_COL else None)
        subtotal_cells.append(f"I{row}")
        row += 2

    # サマリの金額・合計を小計参照にする
    for sr, ref in zip(summary_rows, subtotal_cells):
        c = ws.cell(sr, 9, f"={ref}")
        c.number_format = "#,##0"
        c.font = _f()
    if summary_rows:
        t = ws.cell(total_row, 9, f"=SUM(I{summary_rows[0]}:I{summary_rows[-1]})+N(I{adj_row})")
    else:
        t = ws.cell(total_row, 9, 0)
    t.number_format = "#,##0"
    t.font = _f(bold=True)

    ws.freeze_panes = ws.cell(hr + 1, 1)
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{hr}:{hr}"

    # ---- 物件情報シート ----
    ps = wb.create_sheet("物件情報")
    ps.column_dimensions["A"].width = 22
    ps.column_dimensions["B"].width = 40
    ps.column_dimensions["C"].width = 10
    ps.cell(1, 1, "物件情報（見積条件）").font = _f(14, True)
    rows = [
        ("お客様名", prop.get("customer_name")), ("工事名", prop.get("work_title")),
        ("物件所在地", prop.get("address")), ("構造・工法", prop.get("structure")),
        ("依頼元", prop.get("company_name")), ("提出先（業者）", prop.get("vendor_name")),
        ("見積提出期限", prop.get("due_date")), ("工事希望時期", prop.get("work_period")),
        ("延床面積", _with_unit(prop.get("total_area"), "㎡")), ("延床面積（坪）", _with_unit(prop.get("tsubo"), "坪")),
        ("階数", {1: "平屋", 2: "2階建て", 3: "3階建て"}.get(prop.get("floors"))),
        ("1階床面積", _with_unit(prop.get("floor1_area"), "㎡")), ("2階床面積", _with_unit(prop.get("floor2_area"), "㎡")),
        ("3階床面積", _with_unit(prop.get("floor3_area"), "㎡")),
        ("屋根面積", _with_unit(prop.get("roof_area"), "㎡")), ("外壁面積", _with_unit(prop.get("wall_area"), "㎡")),
        ("足場面積", _with_unit(prop.get("scaffold_area"), "㎡")), ("トイレ数", _with_unit(prop.get("toilet_count"), "ヶ所")),
        ("備考・特記事項", prop.get("remarks")),
    ]
    for r, (k, v) in enumerate(rows, 3):
        ps.cell(r, 1, k).font = _f(bold=True)
        ps.cell(r, 1).fill = HEADER_FILL
        ps.cell(r, 2, v if v not in (None, "") else None).font = _f()
        ps.cell(r, 2).alignment = Alignment(wrap_text=True, vertical="top")
        for c in (1, 2):
            ps.cell(r, c).border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    note_r = len(rows) + 4
    ps.cell(note_r, 1, "※ 黄色のセル（単価・摘要）をご記入ください。金額・小計・合計は自動計算されます。").font = _f()

    _neutralize_formulas(wb)
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def _with_unit(v, unit):
    if v in (None, ""):
        return None
    return f"{_qty(v)} {unit}"


def _neutralize_formulas(wb):
    """利用者の入力(名称・備考など)が「=」で始まると openpyxl は数式として書き込む。
    外部URLを開く数式などを業者側のExcelで実行させないよう、こちらで組んだ金額列(見積書のI列)以外は文字列に戻す。"""
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.data_type == "f" and not (ws.title == "見積書" and cell.column == 9):
                    cell.data_type = "s"
