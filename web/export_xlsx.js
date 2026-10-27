// ---- 単体HTML版: ブラウザ内でExcelを生成して保存する(backend/excel_export.py の移植) ----
// レイアウトを変えるときは excel_export.py とそろえる。
// ExcelJS は文字列を数式として扱わない(数式は { formula } で明示する)ので、入力値はそのまま書いてよい。
const XFONT = "ＭＳ 明朝";
const THIN = { style: "thin" }, HAIR = { style: "hair" };
const HEADER_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9D9D9" } };
const INPUT_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF2CC" } }; // 業者の入力欄(黄色)
const MONEY = "#,##0";

// 列(E・H・J は元フォーマットにある細い余白列)
const COL = { no: 1, name: 2, spec: 3, qty: 4, unit: 6, price: 7, amount: 9, note: 11 };
const LAST_COL = COL.note;
const AMOUNT = "I"; // 金額列(数式で参照する)
const COL_WIDTHS = [6, 38, 34, 9, 2, 7, 14, 2, 16, 2, 28];
const RULED_COLS = Object.values(COL); // 左側に縦罫線を引く列

const xfont = (size = 11, bold = false) => ({ name: XFONT, size, bold });
const nq = v => (v == null || v === "" ? null : Math.round(+v * 100) / 100);

async function buildXlsx({ property: p, categories: cats }) {
  const wb = new ExcelJS.Workbook();
  writeEstimateSheet(wb, p, cats);
  writePropertySheet(wb, p);
  return wb.xlsx.writeBuffer();
}

// 「見積書」シート: ヘッダ → 大項目の一覧 → 調整費 → 合計 → 大項目ごとの内訳と小計
function writeEstimateSheet(wb, p, cats) {
  const ws = wb.addWorksheet("見積書", {
    views: [{ state: "frozen", ySplit: 3 }],
    pageSetup: { paperSize: 9, orientation: "portrait", fitToPage: true, fitToWidth: 1, fitToHeight: 0, printTitlesRow: "3:3" },
  });
  COL_WIDTHS.forEach((w, i) => (ws.getColumn(i + 1).width = w));

  const put = (row, col, value = null, { bold = false, fmt, fill, align } = {}) => {
    const cell = ws.getCell(row, col);
    cell.value = value;
    cell.font = xfont(11, bold);
    if (fmt) cell.numFmt = fmt;
    if (fill) cell.fill = fill;
    if (align) cell.alignment = { horizontal: align };
    return cell;
  };
  // 1行分の罫線。line=THIN で上下を実線にする(ヘッダ・合計・小計)
  const rule = (row, line = HAIR) => {
    for (let c = 1; c <= LAST_COL; c++)
      ws.getCell(row, c).border = {
        left: RULED_COLS.includes(c) ? THIN : undefined,
        right: c === LAST_COL || RULED_COLS.includes(c + 1) ? THIN : undefined,
        top: line, bottom: line,
      };
  };

  // ---- タイトル・宛先 ----
  const title = put(1, COL.name, `${p.customer_name || "お名前"} 様邸${p.work_title || "リノベーション工事"} 見積書`);
  title.font = xfont(14, true); ws.getRow(1).height = 24;
  put(2, COL.name, p.vendor_name ? `${p.vendor_name} 御中` : "");
  ws.mergeCells(2, COL.price, 2, LAST_COL);
  put(2, COL.price, p.company_name || "", { align: "right" });

  // ---- 見出し行 ----
  const headRow = 3;
  rule(headRow, THIN);
  for (let c = 1; c <= LAST_COL; c++) ws.getCell(headRow, c).fill = HEADER_FILL;
  const headings = { [COL.no]: "No.", [COL.name]: "名　　　称", [COL.spec]: "仕　　　様", [COL.qty]: "数量",
    [COL.unit]: "単位", [COL.price]: "単　価", [COL.amount]: "金　額", [COL.note]: "摘　要" };
  for (const [c, text] of Object.entries(headings))
    put(headRow, +c, text, { bold: true }).alignment = { horizontal: "center", vertical: "middle" };
  [COL.qty, COL.price, COL.amount].forEach(c => ws.mergeCells(headRow, c, headRow, c + 1)); // 右隣の余白列と結合

  // ---- 大項目の一覧(金額はあとで各小計への参照を入れる) ----
  let row = headRow + 1;
  put(row++, COL.no, "（内　訳　明　細）");
  const summaryRows = [];
  cats.forEach((cat, i) => {
    put(row, COL.no, i + 1); put(row, COL.name, cat.name); put(row, COL.qty, 1); put(row, COL.unit, "式");
    rule(row); summaryRows.push(row); row++;
  });

  const adjRow = row + 1;
  put(adjRow, COL.name, "調整費");
  put(adjRow, COL.amount, null, { fmt: "#,##0;[Red]-#,##0", fill: INPUT_FILL });
  rule(adjRow);

  const totalRow = adjRow + 2;
  put(totalRow, COL.name, "合　計", { bold: true });
  rule(totalRow, THIN);

  // ---- 大項目ごとの内訳 ----
  row = totalRow + 2;
  const subtotalRows = [];
  cats.forEach((cat, i) => {
    put(row++, COL.no, "（内　訳　明　細）");
    put(row, COL.no, i + 1, { bold: true }); put(row, COL.name, cat.name, { bold: true });
    rule(row); row++;

    const first = row;
    for (const it of cat.items) { writeItemRow(put, row, it); rule(row); row++; }

    put(row, COL.name, "小　計", { bold: true });
    // 項目が0件のとき SUM(I{first}:I{first-1}) は小計セル自身を含む循環参照になるので 0 にする
    const subtotal = row > first ? { formula: `SUM(${AMOUNT}${first}:${AMOUNT}${row - 1})` } : 0;
    put(row, COL.amount, subtotal, { bold: true, fmt: MONEY });
    rule(row, THIN); subtotalRows.push(row); row += 2;
  });

  // 一覧の金額 = 各小計、合計 = 一覧の合計 + 調整費
  summaryRows.forEach((r, i) => put(r, COL.amount, { formula: `${AMOUNT}${subtotalRows[i]}` }, { fmt: MONEY }));
  put(totalRow, COL.amount,
    { formula: `SUM(${AMOUNT}${summaryRows[0]}:${AMOUNT}${summaryRows[summaryRows.length - 1]})+N(${AMOUNT}${adjRow})` },
    { bold: true, fmt: MONEY });
}

// 内訳の1行。見出し(【 】)は名称だけを太字で書く
function writeItemRow(put, row, it) {
  if (it.group) { put(row, COL.name, it.name, { bold: true }); return; }
  put(row, COL.name, it.name);
  put(row, COL.spec, it.spec || null);
  put(row, COL.qty, nq(it.qty), { fmt: "#,##0.##", align: "right" });
  put(row, COL.unit, it.unit || null, { align: "center" });
  put(row, COL.price, null, { fmt: MONEY, fill: INPUT_FILL }); // 単価: 業者が入力
  put(row, COL.amount, { formula: `IF(G${row}="","",ROUND(D${row}*G${row},0))` }, { fmt: MONEY });
  put(row, COL.note, it.note || null, { fill: INPUT_FILL });
}

// 「物件情報」シート
function writePropertySheet(wb, p) {
  const ps = wb.addWorksheet("物件情報");
  ps.getColumn(1).width = 22; ps.getColumn(2).width = 40;
  ps.getCell(1, 1).value = "物件情報（見積条件）"; ps.getCell(1, 1).font = xfont(14, true);
  const u = (v, unit) => (v == null || v === "" ? null : `${nq(v)} ${unit}`);
  const rows = [
    ["お客様名", p.customer_name], ["工事名", p.work_title], ["物件所在地", p.address], ["構造", p.structure],
    ["依頼元", p.company_name], ["提出先（業者）", p.vendor_name], ["見積提出期限", p.due_date], ["工事希望時期", p.work_period],
    ["階数", FLOOR_LABEL[p.floors]],
    ["延床面積", u(p.total_area, "㎡")], ["延床面積（坪）", u(p.tsubo, "坪")],
    ["1階床面積", u(p.floor1_area, "㎡")], ["2階床面積", u(p.floor2_area, "㎡")], ["3階床面積", u(p.floor3_area, "㎡")],
    ["屋根面積", u(p.roof_area, "㎡")], ["外壁面積", u(p.wall_area, "㎡")], ["足場面積", u(p.scaffold_area, "㎡")],
    ["トイレ数", u(p.toilet_count, "ヶ所")], ["備考・特記事項", p.remarks],
  ];
  const box = { left: THIN, right: THIN, top: THIN, bottom: THIN };
  rows.forEach(([label, value], i) => {
    const a = ps.getCell(i + 3, 1), b = ps.getCell(i + 3, 2);
    a.value = label; a.font = xfont(11, true); a.fill = HEADER_FILL; a.border = box;
    b.value = value || null; b.font = xfont(); b.alignment = { wrapText: true, vertical: "top" }; b.border = box;
  });
  const note = ps.getCell(rows.length + 4, 1);
  note.value = "※ 黄色のセル（単価・摘要）をご記入ください。金額・小計・合計は自動計算されます。"; note.font = xfont();
}

// ---- Excelを出力(単体HTML版) ----
$("#export").addEventListener("click", async () => {
  if (!selected.size) return notify("出力する工事を選択してください。");
  const data = exportData(), p = data.property, btn = $("#export");
  btn.disabled = true; btn.textContent = "作成中…";
  try {
    const buf = await buildXlsx(data);
    const filename = [p.customer_name || "お客様", p.work_title, "見積依頼書", p.vendor_name].filter(Boolean).join("_") + ".xlsx";
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
