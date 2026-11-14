"""見積依頼書(Excel)の生成。出力の形は2種類あり、それぞれ見本のExcelの体裁に合わせている。

  renovation(工事)・exterior(外装)
              「リノベーション工事_見積依頼書.xlsx」 … 見積書(8列) + 物件情報
  interior(内装)
              「水回りリフォーム工事 見積書.xlsx」   … 表紙 + 見積もり内訳(11列) + 物件情報

単価は業者が入力する前提のため空欄にし、金額・小計・合計は数式にしてある。
ブラウザ版(web/export_xlsx.js)も同じレイアウトで作っているので、変えるときは両方そろえる。

renovation の「見積書」の並び:
    ヘッダ → 大項目の一覧(金額は各小計を参照) → 調整費 → 合計 → 大項目ごとの内訳と小計
mizumawari の「見積もり内訳」の並び:
    ヘッダ → 大項目ごとの内訳と小計 → 大項目の一覧(金額は各小計を参照) → 合計
"""
from datetime import date
from io import BytesIO

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side

FONT = "ＭＳ 明朝"
THIN = Side(style="thin")
FLOOR_LABELS = {1: "平屋", 2: "2階建て", 3: "3階建て"}
FORMATS = ("renovation", "interior", "exterior")
DEFAULT_TITLE = {"renovation": "リノベーション工事", "interior": "水回りリフォーム工事", "exterior": "外装リフォーム工事"}

SEP = "　"  # 全角スペース(見本の見出しの体裁)
HEAD = {"no": "No.", "name": f"名{SEP*3}称", "spec": f"仕{SEP*3}様", "qty": "数量", "unit": "単位",
        "price": f"単{SEP}価", "amount": f"金{SEP}額", "note": f"摘{SEP}要"}
DETAIL = f"（内{SEP}訳{SEP}明{SEP}細）"


class Layout:
    """シートの列の配置。key: no/name/spec/qty/unit/price/amount/note → 列番号。
    spacer は「直前の列とひと続きに見せる細い余白列」(罫線の左右だけ外す)。"""

    def __init__(self, cols, widths, spacers=()):
        self.cols, self.widths, self.spacers = cols, widths, set(spacers)
        self.last = max(cols.values()) + (1 if (max(cols.values()) + 1) in self.spacers else 0)

    def border(self, c, top=THIN, bottom=THIN):
        left = None if c in self.spacers else THIN
        right = None if (c + 1) in self.spacers else THIN
        return Border(left=left, right=right, top=top, bottom=bottom)


RENOVATION = Layout(
    {"no": 1, "name": 2, "spec": 3, "qty": 4, "unit": 5, "price": 6, "amount": 7, "note": 8},
    [6, 38, 30, 9, 7, 13, 15, 26])
MIZUMAWARI = Layout(
    {"no": 1, "name": 2, "spec": 3, "qty": 4, "unit": 6, "price": 7, "amount": 9, "note": 11},
    [5, 30, 34, 9, 1.33, 7, 12, 1.33, 14, 1.33, 24], spacers=(5, 8, 10))


def _font(size=11, bold=False, name=FONT):
    return Font(name=name, size=size, bold=bold)


def _qty(v):
    """数量の表示用。整数ならintに、それ以外は小数2桁に丸める。"""
    if v is None or v == "":
        return None
    return int(v) if float(v) == int(v) else round(float(v), 2)


def _col(n):
    from openpyxl.utils import get_column_letter
    return get_column_letter(n)


def build_workbook(data: dict) -> BytesIO:
    prop = data["property"]
    fmt = prop.get("format") if prop.get("format") in FORMATS else "renovation"
    wb = Workbook()
    ws = wb.active
    if fmt == "interior":
        ws.title = "表紙"
        est = wb.create_sheet("見積もり内訳")
        total_cell = _write_mizumawari(est, prop, data["categories"], fmt)
        _write_cover(ws, prop, total_cell)
    else:
        ws.title = "見積書"
        _write_renovation(ws, prop, data["categories"], fmt)
    _write_property_sheet(wb.create_sheet("物件情報"), prop)
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ---------------------------------------------------------------- 共通の部品

class Sheet:
    """1枚の見積シートへ書き込む小さな道具。値は文字列のまま書く(「=」で始まる入力が数式にならないよう text 扱い)。"""

    def __init__(self, ws, lay):
        self.ws, self.lay = ws, lay
        for i, w in enumerate(lay.widths, 1):
            ws.column_dimensions[_col(i)].width = w

    def put(self, row, key, value=None, *, bold=False, fmt=None, align=None, formula=None):
        c = self.lay.cols[key] if isinstance(key, str) else key
        cell = self.ws.cell(row, c)
        if formula is not None:
            cell.value = formula
        elif value is not None:
            cell.value = value
            if isinstance(value, str):
                cell.data_type = "s"
        cell.font = _font(bold=bold)
        if fmt:
            cell.number_format = fmt
        cell.alignment = Alignment(horizontal=align or _ALIGN.get(key, "left"), vertical="center")
        return cell

    def grid(self, row, top=THIN, bottom=THIN, height=21):
        """1行分の罫線(見本は全セルが細い実線)と行の高さ。"""
        for c in range(1, self.lay.last + 1):
            self.ws.cell(row, c).border = self.lay.border(c, top, bottom)
        self.ws.row_dimensions[row].height = height

    def detail_header(self, row):
        """「（内 訳 明 細）」の行。全幅を結合する。"""
        self.grid(row)
        self.ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=self.lay.last)
        self.put(row, "no", DETAIL, bold=True, align="left")

    def blank(self, row):
        self.grid(row)

    def item(self, row, it, money):
        """内訳の1行。見出し(【 】)は名称だけを太字で書く。"""
        self.grid(row)
        if it.get("group"):
            self.put(row, "name", it["name"], bold=True)
            return
        a, p, q = _col(self.lay.cols["amount"]), _col(self.lay.cols["price"]), _col(self.lay.cols["qty"])
        self.put(row, "name", it["name"])
        self.put(row, "spec", it.get("spec") or None)
        self.put(row, "qty", _qty(it.get("qty")), align="right")
        self.put(row, "unit", it.get("unit") or None, align="center")
        self.put(row, "price", None, fmt=money, align="right")  # 単価: 業者が入力
        self.put(row, "amount", fmt=money, align="right", formula=f'=IF({p}{row}="","",ROUND({q}{row}*{p}{row},0))')
        self.put(row, "note", it.get("note") or None)

    def head(self, row, title_text, company):
        ws, lay = self.ws, self.lay
        ws.row_dimensions[1].height = 30
        t = ws.cell(1, 2, title_text)
        t.data_type = "s"
        t.font = _font(16, True)
        t.alignment = Alignment(vertical="center")
        ws.row_dimensions[2].height = 21
        if company:
            ws.merge_cells(start_row=2, start_column=lay.cols["price"], end_row=2, end_column=lay.last)
            self.put(2, "price", company, align="right")
        # 見出し行
        self.grid(row, height=21)
        for k, text in HEAD.items():
            self.put(row, k, text, bold=True, align="center")
        for k in ("qty", "price", "amount"):  # 右隣に余白列があれば結合
            c = lay.cols[k]
            if c + 1 in lay.spacers:
                ws.merge_cells(start_row=row, start_column=c, end_row=row, end_column=c + 1)


_ALIGN = {"no": "center", "name": "left", "spec": "left", "qty": "right", "unit": "center",
          "price": "right", "amount": "right", "note": "left"}


def _title(prop, fmt):
    return f"{prop.get('customer_name') or '〇〇'} 様邸{SEP}{prop.get('work_title') or DEFAULT_TITLE[fmt]}{SEP}見積書"


def _page(ws, orientation, last_col, last_row, title_row=None):
    ws.page_setup.orientation = orientation
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.4 if orientation == "landscape" else 0.7
    ws.page_margins.top = 0.5
    ws.print_area = f"A1:{_col(last_col)}{last_row}"
    if title_row:
        ws.print_title_rows = f"{title_row}:{title_row}"


# ---------------------------------------------------------------- リノベーション工事(見積書シート)

def _write_renovation(ws, prop, cats, fmt="renovation"):
    money = "#,##0\\円"
    s = Sheet(ws, RENOVATION)
    s.head(3, _title(prop, fmt), prop.get("company_name"))

    # 大項目の一覧(金額はあとで各小計への参照を入れる)
    s.detail_header(4)
    row, summary_rows = 5, []
    for i, cat in enumerate(cats, 1):
        s.grid(row)
        s.put(row, "no", i)
        s.put(row, "name", cat["name"])
        s.put(row, "qty", 1, align="right")
        s.put(row, "unit", "式", align="center")
        summary_rows.append(row)
        row += 1
    s.blank(row)
    adj_row = row + 1
    s.grid(adj_row)
    s.put(adj_row, "name", "調整費")
    s.put(adj_row, "amount", fmt=money, align="right")
    s.blank(adj_row + 1)
    total_row = adj_row + 2
    s.grid(total_row)
    s.put(total_row, "name", f"合{SEP}計", bold=True)

    # 大項目ごとの内訳
    row = total_row + 2
    subtotal_rows = []
    for i, cat in enumerate(cats, 1):
        s.detail_header(row)
        row += 1
        s.grid(row)
        s.put(row, "no", i, bold=True)
        s.put(row, "name", cat["name"], bold=True)
        row += 1
        first = row
        for it in cat["items"]:
            s.item(row, it, money)
            row += 1
        s.grid(row)
        s.put(row, "name", f"小{SEP}計", bold=True)
        # 項目が0件のとき SUM(G{first}:G{first-1}) は小計セル自身を含む循環参照になるので 0 にする
        s.put(row, "amount", 0 if row == first else None, bold=True, fmt=money, align="right",
              formula=None if row == first else f"=SUM(G{first}:G{row - 1})")
        subtotal_rows.append(row)
        row += 2

    for sr, tr in zip(summary_rows, subtotal_rows):
        s.put(sr, "amount", fmt=money, align="right", formula=f"=G{tr}")
    total = f"=SUM(G{summary_rows[0]}:G{summary_rows[-1]})+N(G{adj_row})" if summary_rows else None
    s.put(total_row, "amount", 0 if total is None else None, bold=True, fmt=money, align="right", formula=total)

    ws.freeze_panes = "A4"
    _page(ws, "portrait", RENOVATION.last, row - 2, 3)


# ---------------------------------------------------------------- 水回りリフォーム(表紙 + 見積もり内訳)

def _write_mizumawari(ws, prop, cats, fmt="interior"):
    """見積もり内訳シートを書き、合計セルの番地(表紙から参照する)を返す。"""
    money = "#,##0_);[Red](#,##0)"
    s = Sheet(ws, MIZUMAWARI)
    for col in range(12, 27):  # 見本は L 列以降にも幅を持たせている
        ws.column_dimensions[_col(col)].width = 14.5 if col == 12 else 8
    s.head(3, _title(prop, fmt), prop.get("company_name"))
    s.detail_header(4)

    row, subtotal_rows = 5, []
    for i, cat in enumerate(cats, 1):
        s.grid(row)
        s.put(row, "no", i, bold=True)
        s.put(row, "name", cat["name"], bold=True)
        row += 1
        first = row
        for it in cat["items"]:
            s.item(row, it, money)
            row += 1
        s.blank(row)  # 業者が項目を書き足せる予備の行(小計に含まれる)
        row += 1
        s.grid(row)
        s.put(row, "name", f"小{SEP}計", bold=True, align="center")
        s.put(row, "amount", bold=True, fmt=money, align="right", formula=f"=SUM(I{first}:I{row - 1})")
        subtotal_rows.append(row)
        row += 1
        s.blank(row)
        row += 1

    # 大項目の一覧と合計
    row += 1
    s.detail_header(row)
    row += 1
    first_sum = row
    for i, (cat, tr) in enumerate(zip(cats, subtotal_rows), 1):
        s.grid(row)
        s.put(row, "no", i)
        s.put(row, "name", cat["name"])
        s.put(row, "qty", 1, align="right")
        s.put(row, "unit", "式", align="center")
        s.put(row, "amount", fmt=money, align="right", formula=f"=I{tr}")
        row += 1
    s.blank(row)  # 予備の行(諸経費などを書き足せる)
    row += 1
    s.grid(row)
    s.put(row, "name", f"合{SEP}計", bold=True, align="center")
    s.put(row, "amount", bold=True, fmt=money, align="right", formula=f"=SUM(I{first_sum}:I{row - 1})")
    total_row = row

    ws.freeze_panes = "A4"
    _page(ws, "landscape", MIZUMAWARI.last, total_row, 3)
    return f"I{total_row}"


COVER_COL_WIDTHS = [3.66] * 20 + [5.16] + [3.66] * 9 + [3.5] * 4 + [3.66] * 3 + [9]
COVER_ROW_HEIGHTS = {1: 18, 2: 18, 3: 39, 4: 11.25, 5: 18, 6: 19.5, 7: 31.5, 8: 12.75, 9: 15.75, 10: 18, 11: 18, 12: 21.75,
                     13: 19.5, 14: 12, 15: 19.5, 16: 22.5, 17: 18, 18: 18, 19: 18, 20: 18, 21: 15, 22: 18, 23: 15, 24: 18,
                     25: 15, 26: 18, 27: 15, 28: 18, 29: 14.25, 30: 18, 31: 18}


def _write_cover(ws, prop, total_cell):
    """表紙。御見積金額は見積もり内訳の合計を参照する(単価が入ると金額が出る)。"""
    for i, w in enumerate(COVER_COL_WIDTHS, 1):
        ws.column_dimensions[_col(i)].width = w
    for r, h in COVER_ROW_HEIGHTS.items():
        ws.row_dimensions[r].height = h

    def put(addr, value=None, size=12, bold=False, h="left", v=None, fmt=None, merge=None, formula=None, name=FONT):
        if merge:
            ws.merge_cells(merge)
        c = ws[addr]
        if formula is not None:
            c.value = formula
        elif value is not None:
            c.value = value
            if isinstance(value, str):
                c.data_type = "s"
        c.font = _font(size, bold, name)
        c.alignment = Alignment(horizontal=h, vertical=v, wrap_text=True if v == "top" else None)
        if fmt:
            c.number_format = fmt
        return c

    put("A1", f"御{SEP}見{SEP}積{SEP}書", 24, True, "center", merge="A1:AJ3")
    today = date.today()
    put("AC5", f"令和{today.year - 2018}", h="right", v="center", merge="AC5:AD5")
    put("AE5", "年", v="center")
    put("AF5", today.month, h="right", v="center")
    put("AG5", "月", v="center")
    put("AH5", today.day, h="right", v="center")
    put("AI5", "日", v="center")
    put("B7", f"{prop.get('customer_name') or ''}{SEP}様", 18, merge="B7:N7")

    # 御見積金額 = 工事金額 + 消費税相当額
    put("C10", "御見積金額 ", 22, h="center", merge="C10:J11")
    put("K10", "￥", 22, h="center", merge="K10:K11")
    put("L10", None, 22, h="center", merge="L10:U11", fmt="[DBNum3][$-411]#,##0", formula="=H13+S13")
    put("V10", "円", 22, h="center", merge="V10:W11")
    for col in range(3, 24):  # 金額欄の下線
        ws.cell(11, col).border = Border(bottom=THIN)
    put("C13", "内訳")
    put("E13", "工事金額 ")
    put("H13", None, 14, h="center", merge="H13:L13", fmt='"¥"#,##0;[Red]"¥"\\-#,##0', formula=f"='見積もり内訳'!{total_cell}")
    put("O13", "消費税相当額", h="center", merge="O13:R13")
    put("S13", None, 14, h="center", merge="S13:V13", fmt='"¥"#,##0;[Red]"¥"\\-#,##0', formula="=H13*0.1")
    put("H15", " 上記の通り御見積申し上げます。", 14, h="center", merge="H15:S15")

    # 検印欄
    for left, label in (("Z", "検印"), ("AC", "検印"), ("AF", "作成者")):
        li = ws[f"{left}1"].column
        put(f"{left}18", label, 11, h="center", merge=f"{left}18:{_col(li + 2)}18")
        ws.merge_cells(f"{left}19:{_col(li + 2)}21")
        for r in range(18, 22):
            for c in range(li, li + 3):
                ws.cell(r, c).border = Border(
                    left=THIN if c == li else None, right=THIN if c == li + 2 else None,
                    top=THIN if r in (18, 19) else None, bottom=THIN if r in (18, 21) else None)

    # 工事名・場所など
    remarks = prop.get("remarks") or ""
    rows = [(20, "工事名", prop.get("work_title") or DEFAULT_TITLE["interior"], "F20:U20"),
            (22, "工事場所", prop.get("address"), "F22:O22"),
            (24, "工期", prop.get("work_period"), "F24:O24"),
            (26, "決済条件", None, "F26:O26"),
            (28, "有効期限", None, "F28:O28"),
            (30, "備考", remarks or None, "F30:O30")]
    for r, label, value, rng in rows:
        put(f"B{r}", label, h="center", v="center", merge=f"B{r}:D{r}")
        put(f"F{r}", value, h="left", v="center", merge=rng)
    if len(remarks) > 24:  # 備考が長いときは行を高くして折り返す
        ws.row_dimensions[30].height = min(18 * (len(remarks) // 24 + 1), 160)
        ws["F30"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

    # 依頼元(自社名)
    if prop.get("company_name"):
        put("Y24", prop["company_name"], 18, True, "center", "center", merge="Y24:AJ25", name="HG丸ｺﾞｼｯｸM-PRO")

    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = 0.5
    ws.print_area = "A1:AJ31"


# ---------------------------------------------------------------- 物件情報シート

def _write_property_sheet(ps, prop):
    ps.column_dimensions["A"].width = 22
    ps.column_dimensions["B"].width = 48
    ps.cell(1, 1, "物件情報（見積条件）").font = _font(16, True)
    rows = [
        ("お客様名", prop.get("customer_name")),
        ("工事名", prop.get("work_title")),
        ("物件所在地", prop.get("address")),
        ("構造・工法", prop.get("structure")),
        ("依頼元", prop.get("company_name")),
        ("見積提出期限", prop.get("due_date")),
        ("工事希望時期", prop.get("work_period")),
        ("階数", FLOOR_LABELS.get(prop.get("floors"))),
        ("水回りの範囲", prop.get("plumbing")),
        ("備考・特記事項", prop.get("remarks")),
    ]
    box = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    for r, (label, value) in enumerate(rows, 3):
        a = ps.cell(r, 1, label)
        a.font = _font(bold=True)
        b = ps.cell(r, 2, value if value not in (None, "") else None)
        if isinstance(b.value, str):
            b.data_type = "s"
        b.font = _font()
        b.alignment = Alignment(wrap_text=True, vertical="center" if r != 12 else "top")
        a.alignment = Alignment(vertical="center")
        a.border = b.border = box
        ps.row_dimensions[r].height = 21
    ps.row_dimensions[12].height = 63  # 備考
    ps.cell(len(rows) + 4, 1, "※ 単価・摘要をご記入ください。金額・小計・合計は自動計算されます。").font = _font(10)
