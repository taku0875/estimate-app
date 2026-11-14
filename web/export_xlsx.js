// ---- 単体HTML版: ブラウザ内でExcelを生成して保存する(backend/excel_export.py の移植) ----
// レイアウトを変えるときは excel_export.py とそろえる。出力の形は2種類(見本のExcelに合わせている):
//   renovation(工事)・exterior(外装) … 見積書(8列) + 物件情報 / interior(内装) … 表紙 + 見積もり内訳(11列) + 物件情報
// ExcelJS は文字列を数式として扱わない(数式は { formula } で明示する)ので、入力値はそのまま書いてよい。
const XFONT = "ＭＳ 明朝";
const THIN = { style: "thin" };
const SEP = "　";
const HEAD = { no: "No.", name: `名${SEP}${SEP}${SEP}称`, spec: `仕${SEP}${SEP}${SEP}様`, qty: "数量", unit: "単位", price: `単${SEP}価`, amount: `金${SEP}額`, note: `摘${SEP}要` };
const DETAIL = `（内${SEP}訳${SEP}明${SEP}細）`;
const DEFAULT_TITLE = { renovation: "リノベーション工事", interior: "水回りリフォーム工事", exterior: "外装リフォーム工事" };
const ALIGN = { no: "center", name: "left", spec: "left", qty: "right", unit: "center", price: "right", amount: "right", note: "left" };

// 列の配置。spacers は「直前の列とひと続きに見せる細い余白列」(罫線の左右だけ外す)
const LAYOUTS = {
  renovation: { cols: { no: 1, name: 2, spec: 3, qty: 4, unit: 5, price: 6, amount: 7, note: 8 }, widths: [6, 38, 30, 9, 7, 13, 15, 26], spacers: [], last: 8 },
  mizumawari: { cols: { no: 1, name: 2, spec: 3, qty: 4, unit: 6, price: 7, amount: 9, note: 11 }, widths: [5, 30, 34, 9, 1.33, 7, 12, 1.33, 14, 1.33, 24], spacers: [5, 8, 10], last: 11 },
};

const xfont = (size = 11, bold = false, name = XFONT) => ({ name, size, bold });
const nq = v => (v == null || v === "" ? null : Math.round(+v * 100) / 100);
const colL = n => { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };

async function buildXlsx({ property: p, categories: cats }) {
  const wb = new ExcelJS.Workbook();
  if (p.format === "interior") {
    const cover = wb.addWorksheet("表紙");
    const est = wb.addWorksheet("見積もり内訳");
    const totalCell = writeMizumawari(est, p, cats, p.format);
    writeCover(cover, p, totalCell);
  } else writeRenovation(wb.addWorksheet("見積書"), p, cats, p.format);
  writePropertySheet(wb.addWorksheet("物件情報"), p);
  return wb.xlsx.writeBuffer();
}

// 見積シートへ書き込む小さな道具
function makeSheet(ws, lay) {
  lay.widths.forEach((w, i) => (ws.getColumn(i + 1).width = w));
  const put = (row, key, value = null, { bold = false, fmt, align, formula } = {}) => {
    const c = typeof key === "string" ? lay.cols[key] : key;
    const cell = ws.getCell(row, c);
    cell.value = formula != null ? { formula } : value;
    cell.font = xfont(11, bold);
    if (fmt) cell.numFmt = fmt;
    cell.alignment = { horizontal: align || ALIGN[key] || "left", vertical: "middle" };
    return cell;
  };
  // 1行分の罫線(見本は全セルが細い実線)と行の高さ
  const grid = (row, height = 21) => {
    for (let c = 1; c <= lay.last; c++)
      ws.getCell(row, c).border = {
        top: THIN, bottom: THIN,
        left: lay.spacers.includes(c) ? undefined : THIN,
        right: lay.spacers.includes(c + 1) ? undefined : THIN,
      };
    ws.getRow(row).height = height;
  };
  const detailHeader = row => { grid(row); ws.mergeCells(row, 1, row, lay.last); put(row, "no", DETAIL, { bold: true, align: "left" }); };
  const item = (row, it, money) => {
    grid(row);
    if (it.group) { put(row, "name", it.name, { bold: true }); return; }
    const p = colL(lay.cols.price), q = colL(lay.cols.qty);
    put(row, "name", it.name);
    put(row, "spec", it.spec || null);
    put(row, "qty", nq(it.qty), { align: "right" });
    put(row, "unit", it.unit || null, { align: "center" });
    put(row, "price", null, { fmt: money, align: "right" }); // 単価: 業者が入力
    put(row, "amount", null, { fmt: money, align: "right", formula: `IF(${p}${row}="","",ROUND(${q}${row}*${p}${row},0))` });
    put(row, "note", it.note || null);
  };
  const head = (row, titleText, company) => {
    ws.getRow(1).height = 30;
    const t = ws.getCell(1, 2); t.value = titleText; t.font = xfont(16, true); t.alignment = { vertical: "middle" };
    ws.getRow(2).height = 21;
    if (company) { ws.mergeCells(2, lay.cols.price, 2, lay.last); put(2, "price", company, { align: "right" }); }
    grid(row);
    for (const [k, text] of Object.entries(HEAD)) put(row, k, text, { bold: true, align: "center" });
    for (const k of ["qty", "price", "amount"]) if (lay.spacers.includes(lay.cols[k] + 1)) ws.mergeCells(row, lay.cols[k], row, lay.cols[k] + 1);
  };
  return { put, grid, detailHeader, item, head };
}

const titleOf = (p, fmt) => `${p.customer_name || "〇〇"} 様邸${SEP}${p.work_title || DEFAULT_TITLE[fmt]}${SEP}見積書`;

function pageSetup(ws, orientation, lastCol, lastRow, titleRow) {
  const m = orientation === "landscape" ? 0.4 : 0.7;
  ws.pageSetup = {
    paperSize: 9, orientation, fitToPage: true, fitToWidth: 1, fitToHeight: 0,
    horizontalCentered: true, margins: { left: m, right: m, top: 0.5, bottom: 0.75, header: 0.3, footer: 0.3 },
    printArea: `A1:${colL(lastCol)}${lastRow}`, printTitlesRow: `${titleRow}:${titleRow}`,
  };
}

// リノベーション工事「見積書」シート: ヘッダ → 大項目の一覧 → 調整費 → 合計 → 大項目ごとの内訳と小計
function writeRenovation(ws, p, cats, fmt = "renovation") {
  const money = "#,##0\\円", lay = LAYOUTS.renovation, s = makeSheet(ws, lay);
  s.head(3, titleOf(p, fmt), p.company_name);

  s.detailHeader(4);
  let row = 5; const summaryRows = [];
  cats.forEach((cat, i) => {
    s.grid(row); s.put(row, "no", i + 1); s.put(row, "name", cat.name); s.put(row, "qty", 1, { align: "right" }); s.put(row, "unit", "式", { align: "center" });
    summaryRows.push(row); row++;
  });
  s.grid(row);
  const adjRow = row + 1; s.grid(adjRow); s.put(adjRow, "name", "調整費"); s.put(adjRow, "amount", null, { fmt: money, align: "right" });
  s.grid(adjRow + 1);
  const totalRow = adjRow + 2; s.grid(totalRow); s.put(totalRow, "name", `合${SEP}計`, { bold: true });

  row = totalRow + 2; const subtotalRows = [];
  cats.forEach((cat, i) => {
    s.detailHeader(row++);
    s.grid(row); s.put(row, "no", i + 1, { bold: true }); s.put(row, "name", cat.name, { bold: true }); row++;
    const first = row;
    for (const it of cat.items) s.item(row++, it, money);
    s.grid(row); s.put(row, "name", `小${SEP}計`, { bold: true });
    // 項目が0件のとき SUM(G{first}:G{first-1}) は小計セル自身を含む循環参照になるので 0 にする
    if (row > first) s.put(row, "amount", null, { bold: true, fmt: money, align: "right", formula: `SUM(G${first}:G${row - 1})` });
    else s.put(row, "amount", 0, { bold: true, fmt: money, align: "right" });
    subtotalRows.push(row); row += 2;
  });

  summaryRows.forEach((r, i) => s.put(r, "amount", null, { fmt: money, align: "right", formula: `G${subtotalRows[i]}` }));
  if (summaryRows.length) s.put(totalRow, "amount", null, { bold: true, fmt: money, align: "right", formula: `SUM(G${summaryRows[0]}:G${summaryRows[summaryRows.length - 1]})+N(G${adjRow})` });
  ws.views = [{ state: "frozen", ySplit: 3 }];
  pageSetup(ws, "portrait", lay.last, row - 2, 3);
}

// 水回りリフォーム「見積もり内訳」シート: ヘッダ → 大項目ごとの内訳と小計 → 大項目の一覧 → 合計。合計セルの番地を返す
function writeMizumawari(ws, p, cats, fmt = "interior") {
  const money = "#,##0_);[Red](#,##0)", lay = LAYOUTS.mizumawari, s = makeSheet(ws, lay);
  ws.getColumn(12).width = 14.5; for (let c = 13; c <= 26; c++) ws.getColumn(c).width = 8;
  s.head(3, titleOf(p, fmt), p.company_name);
  s.detailHeader(4);

  let row = 5; const subtotalRows = [];
  cats.forEach((cat, i) => {
    s.grid(row); s.put(row, "no", i + 1, { bold: true }); s.put(row, "name", cat.name, { bold: true }); row++;
    const first = row;
    for (const it of cat.items) s.item(row++, it, money);
    s.grid(row++); // 業者が項目を書き足せる予備の行(小計に含まれる)
    s.grid(row); s.put(row, "name", `小${SEP}計`, { bold: true, align: "center" });
    s.put(row, "amount", null, { bold: true, fmt: money, align: "right", formula: `SUM(I${first}:I${row - 1})` });
    subtotalRows.push(row++);
    s.grid(row++);
  });

  row++;
  s.detailHeader(row++);
  const firstSum = row;
  cats.forEach((cat, i) => {
    s.grid(row); s.put(row, "no", i + 1); s.put(row, "name", cat.name); s.put(row, "qty", 1, { align: "right" }); s.put(row, "unit", "式", { align: "center" });
    s.put(row, "amount", null, { fmt: money, align: "right", formula: `I${subtotalRows[i]}` });
    row++;
  });
  s.grid(row++); // 予備の行(諸経費などを書き足せる)
  s.grid(row); s.put(row, "name", `合${SEP}計`, { bold: true, align: "center" });
  s.put(row, "amount", null, { bold: true, fmt: money, align: "right", formula: `SUM(I${firstSum}:I${row - 1})` });
  ws.views = [{ state: "frozen", ySplit: 3 }];
  pageSetup(ws, "landscape", lay.last, row, 3);
  return `I${row}`;
}

// 表紙。御見積金額は見積もり内訳の合計を参照する(単価が入ると金額が出る)
const COVER_WIDTHS = [...Array(20).fill(3.66), 5.16, ...Array(9).fill(3.66), ...Array(4).fill(3.5), ...Array(3).fill(3.66), 9];
const COVER_HEIGHTS = { 1: 18, 2: 18, 3: 39, 4: 11.25, 5: 18, 6: 19.5, 7: 31.5, 8: 12.75, 9: 15.75, 10: 18, 11: 18, 12: 21.75, 13: 19.5, 14: 12, 15: 19.5, 16: 22.5, 17: 18, 18: 18, 19: 18, 20: 18, 21: 15, 22: 18, 23: 15, 24: 18, 25: 15, 26: 18, 27: 15, 28: 18, 29: 14.25, 30: 18, 31: 18 };
function writeCover(ws, p, totalCell) {
  COVER_WIDTHS.forEach((w, i) => (ws.getColumn(i + 1).width = w));
  Object.entries(COVER_HEIGHTS).forEach(([r, h]) => (ws.getRow(+r).height = h));
  const put = (addr, value = null, { size = 12, bold = false, h = "left", v, fmt, merge, formula, name = XFONT } = {}) => {
    if (merge) ws.mergeCells(merge);
    const c = ws.getCell(addr);
    c.value = formula != null ? { formula } : value;
    c.font = xfont(size, bold, name);
    c.alignment = { horizontal: h, vertical: v };
    if (fmt) c.numFmt = fmt;
    return c;
  };
  put("A1", `御${SEP}見${SEP}積${SEP}書`, { size: 24, bold: true, h: "center", merge: "A1:AJ3" });
  const now = new Date();
  put("AC5", `令和${now.getFullYear() - 2018}`, { h: "right", v: "middle", merge: "AC5:AD5" });
  put("AE5", "年", { v: "middle" }); put("AF5", now.getMonth() + 1, { h: "right", v: "middle" });
  put("AG5", "月", { v: "middle" }); put("AH5", now.getDate(), { h: "right", v: "middle" }); put("AI5", "日", { v: "middle" });
  put("B7", `${p.customer_name || ""}${SEP}様`, { size: 18, merge: "B7:N7" });

  put("C10", "御見積金額 ", { size: 22, h: "center", merge: "C10:J11" });
  put("K10", "￥", { size: 22, h: "center", merge: "K10:K11" });
  put("L10", null, { size: 22, h: "center", merge: "L10:U11", fmt: "[DBNum3][$-411]#,##0", formula: "H13+S13" });
  put("V10", "円", { size: 22, h: "center", merge: "V10:W11" });
  for (let c = 3; c <= 23; c++) ws.getCell(11, c).border = { bottom: THIN };
  put("C13", "内訳"); put("E13", "工事金額 ");
  const yen = "\"¥\"#,##0;[Red]\"¥\"\\-#,##0";
  put("H13", null, { size: 14, h: "center", merge: "H13:L13", fmt: yen, formula: `'見積もり内訳'!${totalCell}` });
  put("O13", "消費税相当額", { h: "center", merge: "O13:R13" });
  put("S13", null, { size: 14, h: "center", merge: "S13:V13", fmt: yen, formula: "H13*0.1" });
  put("H15", " 上記の通り御見積申し上げます。", { size: 14, h: "center", merge: "H15:S15" });

  // 検印欄
  [[26, "検印"], [29, "検印"], [32, "作成者"]].forEach(([li, label]) => {
    const a = colL(li), b = colL(li + 2);
    put(`${a}18`, label, { size: 11, h: "center", merge: `${a}18:${b}18` });
    ws.mergeCells(`${a}19:${b}21`);
    for (let r = 18; r <= 21; r++) for (let c = li; c <= li + 2; c++)
      ws.getCell(r, c).border = { left: c === li ? THIN : undefined, right: c === li + 2 ? THIN : undefined, top: r <= 19 ? THIN : undefined, bottom: r === 18 || r === 21 ? THIN : undefined };
  });

  const remarks = p.remarks || "";
  [[20, "工事名", p.work_title || DEFAULT_TITLE.interior, "F20:U20"], [22, "工事場所", p.address, "F22:O22"], [24, "工期", p.work_period, "F24:O24"],
   [26, "決済条件", null, "F26:O26"], [28, "有効期限", null, "F28:O28"], [30, "備考", remarks || null, "F30:O30"]].forEach(([r, label, value, rng]) => {
    put(`B${r}`, label, { h: "center", v: "middle", merge: `B${r}:D${r}` });
    put(`F${r}`, value || null, { h: "left", v: "middle", merge: rng });
  });
  if (remarks.length > 24) { // 備考が長いときは行を高くして折り返す
    ws.getRow(30).height = Math.min(18 * (Math.floor(remarks.length / 24) + 1), 160);
    ws.getCell("F30").alignment = { horizontal: "left", vertical: "top", wrapText: true };
  }
  if (p.company_name) put("Y24", p.company_name, { size: 18, bold: true, h: "center", v: "middle", merge: "Y24:AJ25", name: "HG丸ｺﾞｼｯｸM-PRO" });

  ws.pageSetup = { paperSize: 9, orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 1, horizontalCentered: true,
    margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.75, header: 0.3, footer: 0.3 }, printArea: "A1:AJ31" };
}

// 「物件情報」シート
function writePropertySheet(ps, p) {
  ps.getColumn(1).width = 22; ps.getColumn(2).width = 48;
  ps.getCell(1, 1).value = "物件情報（見積条件）"; ps.getCell(1, 1).font = xfont(16, true);
  const rows = [
    ["お客様名", p.customer_name], ["工事名", p.work_title], ["物件所在地", p.address], ["構造・工法", p.structure],
    ["依頼元", p.company_name], ["見積提出期限", p.due_date], ["工事希望時期", p.work_period],
    ["階数", FLOOR_LABEL[p.floors]], ["水回りの範囲", p.plumbing], ["備考・特記事項", p.remarks],
  ];
  const box = { left: THIN, right: THIN, top: THIN, bottom: THIN };
  rows.forEach(([label, value], i) => {
    const r = i + 3, a = ps.getCell(r, 1), b = ps.getCell(r, 2);
    a.value = label; a.font = xfont(11, true); a.border = box; a.alignment = { vertical: "middle" };
    b.value = value || null; b.font = xfont(); b.alignment = { wrapText: true, vertical: r === 12 ? "top" : "middle" }; b.border = box;
    ps.getRow(r).height = r === 12 ? 63 : 21;
  });
  const note = ps.getCell(rows.length + 4, 1);
  note.value = "※ 単価・摘要をご記入ください。金額・小計・合計は自動計算されます。"; note.font = xfont(10);
}

// ---- Excelを出力(単体HTML版) ----
$("#export").addEventListener("click", async () => {
  if (!selected.size) return notify("出力する工事を選択してください。");
  const data = exportData(), p = data.property, btn = $("#export");
  btn.disabled = true; btn.textContent = "作成中…";
  try {
    const buf = await buildXlsx(data);
    const filename = [p.customer_name || "お客様", p.work_title, "見積依頼書"].filter(Boolean).join("_") + ".xlsx";
    // claude.ai 上の公開ページでは保存機能(downloads)を使い、通常のWebサイトでは通常のダウンロードにする
    const downloads = window.claude && claude.use ? await claude.use("downloads").catch(() => null) : null;
    if (downloads) await downloads.save({ filename, data: buf });
    else {
      const a = Object.assign(document.createElement("a"), {
        href: URL.createObjectURL(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" })), download: filename,
      });
      document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 4000);
    }
    notify("Excelを出力しました。");
  } catch (err) {
    if (err && err.code === "declined") notify("保存をキャンセルしました。");
    else notify("出力に失敗しました: " + (err && err.message ? err.message : err));
  } finally { btn.disabled = false; btn.textContent = "Excelを出力"; }
});

// 下部のステータス欄に一時的にメッセージを出す
function notify(msg) {
  $("#status").textContent = msg;
  setTimeout(() => renderDetail(), 3500);
}
