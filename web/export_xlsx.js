// ---- サンプル版: ブラウザ内でExcelを生成して保存する(backend/excel_export.py の移植) ----
const XFONT = "ＭＳ 明朝";
const xf = (size = 11, bold = false) => ({ name: XFONT, size, bold });
const THIN = { style: "thin" }, HAIR = { style: "hair" };
const GREY = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9D9D9" } };
const YELLOW = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF2CC" } };
const LINES = [1, 2, 3, 4, 6, 7, 9, 11];
const nq = v => (v == null || v === "" ? null : Math.round(+v * 100) / 100);

async function buildXlsx({ property: p, categories: cats }) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("見積書", {
    views: [{ state: "frozen", ySplit: 3 }],
    pageSetup: { paperSize: 9, orientation: "portrait", fitToPage: true, fitToWidth: 1, fitToHeight: 0, printTitlesRow: "3:3" },
  });
  [6, 38, 34, 9, 2, 7, 14, 2, 16, 2, 28].forEach((w, i) => (ws.getColumn(i + 1).width = w));
  const cell = (r, c) => ws.getCell(r, c);
  const frame = (r, top, bottom) => {
    for (let c = 1; c <= 11; c++)
      cell(r, c).border = {
        left: LINES.includes(c) ? THIN : undefined, right: c === 11 || [1, 2, 3, 5, 6, 8, 10].includes(c) ? THIN : undefined,
        top: top || HAIR, bottom: bottom || HAIR,
      };
  };

  cell(1, 2).value = `${p.customer_name || "お名前"} 様邸${p.work_title || "リノベーション工事"} 見積書`;
  cell(1, 2).font = xf(14, true); ws.getRow(1).height = 24;
  cell(2, 2).value = p.vendor_name ? `${p.vendor_name} 御中` : ""; cell(2, 2).font = xf();
  ws.mergeCells(2, 7, 2, 11);
  cell(2, 7).value = p.company_name || ""; cell(2, 7).font = xf(); cell(2, 7).alignment = { horizontal: "right" };

  const hr = 3;
  frame(hr, THIN, THIN);
  for (let c = 1; c <= 11; c++) cell(hr, c).fill = GREY;
  Object.entries({ 1: "No.", 2: "名　　　称", 3: "仕　　　様", 4: "数量", 6: "単位", 7: "単　価", 9: "金　額", 11: "摘　要" }).forEach(([c, t]) => {
    cell(hr, +c).value = t; cell(hr, +c).font = xf(11, true); cell(hr, +c).alignment = { horizontal: "center", vertical: "middle" };
  });
  [[4, 5], [7, 8], [9, 10]].forEach(([a, b]) => ws.mergeCells(hr, a, hr, b));

  let row = hr + 1;
  cell(row++, 1).value = "（内　訳　明　細）"; cell(row - 1, 1).font = xf();
  const summaryRows = [];
  cats.forEach((cat, i) => {
    cell(row, 1).value = i + 1; cell(row, 2).value = cat.name; cell(row, 4).value = 1; cell(row, 6).value = "式";
    [1, 2, 4, 6].forEach(c => (cell(row, c).font = xf()));
    frame(row); summaryRows.push(row); row++;
  });
  const adj = row + 1;
  cell(adj, 2).value = "調整費"; cell(adj, 2).font = xf();
  cell(adj, 9).fill = YELLOW; cell(adj, 9).numFmt = "#,##0;[Red]-#,##0"; frame(adj);
  const total = adj + 2;
  cell(total, 2).value = "合　計"; cell(total, 2).font = xf(11, true); frame(total, THIN, THIN);
  row = total + 2;

  const subs = [];
  cats.forEach((cat, i) => {
    cell(row++, 1).value = "（内　訳　明　細）"; cell(row - 1, 1).font = xf();
    cell(row, 1).value = i + 1; cell(row, 2).value = cat.name;
    cell(row, 1).font = xf(11, true); cell(row, 2).font = xf(11, true); frame(row); row++;
    const first = row;
    for (const it of cat.items) {
      cell(row, 2).value = it.name; cell(row, 2).font = xf(11, !!it.group);
      if (!it.group) {
        cell(row, 3).value = it.spec || null; cell(row, 3).font = xf();
        const q = cell(row, 4); q.value = nq(it.qty); q.font = xf(); q.numFmt = "#,##0.##"; q.alignment = { horizontal: "right" };
        cell(row, 6).value = it.unit || null; cell(row, 6).font = xf(); cell(row, 6).alignment = { horizontal: "center" };
        const g = cell(row, 7); g.fill = YELLOW; g.numFmt = "#,##0"; g.font = xf();
        const a = cell(row, 9); a.value = { formula: `IF(G${row}="","",ROUND(D${row}*G${row},0))` }; a.numFmt = "#,##0"; a.font = xf();
        cell(row, 11).value = it.note || null; cell(row, 11).font = xf(); cell(row, 11).fill = YELLOW;
      }
      frame(row); row++;
    }
    cell(row, 2).value = "小　計"; cell(row, 2).font = xf(11, true);
    const s = cell(row, 9); s.value = { formula: `SUM(I${first}:I${row - 1})` }; s.numFmt = "#,##0"; s.font = xf(11, true);
    frame(row, THIN, THIN); subs.push(`I${row}`); row += 2;
  });

  summaryRows.forEach((r, i) => { const c = cell(r, 9); c.value = { formula: subs[i] }; c.numFmt = "#,##0"; c.font = xf(); });
  const t = cell(total, 9);
  t.value = { formula: `SUM(I${summaryRows[0]}:I${summaryRows[summaryRows.length - 1]})+N(I${adj})` }; t.numFmt = "#,##0"; t.font = xf(11, true);

  const ps = wb.addWorksheet("物件情報");
  ps.getColumn(1).width = 22; ps.getColumn(2).width = 40;
  ps.getCell(1, 1).value = "物件情報（見積条件）"; ps.getCell(1, 1).font = xf(14, true);
  const u = (v, unit) => (v == null || v === "" ? null : `${nq(v)} ${unit}`);
  const rows = [
    ["お客様名", p.customer_name], ["工事名", p.work_title], ["物件所在地", p.address], ["構造", p.structure],
    ["依頼元", p.company_name], ["提出先（業者）", p.vendor_name], ["見積提出期限", p.due_date], ["工事希望時期", p.work_period],
    ["階数", { 1: "平屋", 2: "2階建て", 3: "3階建て" }[p.floors]],
    ["延床面積", u(p.total_area, "㎡")], ["延床面積（坪）", u(p.tsubo, "坪")],
    ["1階床面積", u(p.floor1_area, "㎡")], ["2階床面積", u(p.floor2_area, "㎡")], ["3階床面積", u(p.floor3_area, "㎡")],
    ["屋根面積", u(p.roof_area, "㎡")], ["外壁面積", u(p.wall_area, "㎡")], ["足場面積", u(p.scaffold_area, "㎡")],
    ["トイレ数", u(p.toilet_count, "ヶ所")], ["備考・特記事項", p.remarks],
  ];
  rows.forEach(([k, v], i) => {
    const r = i + 3, a = ps.getCell(r, 1), b = ps.getCell(r, 2);
    a.value = k; a.font = xf(11, true); a.fill = GREY; b.value = v || null; b.font = xf();
    b.alignment = { wrapText: true, vertical: "top" };
    [a, b].forEach(c => (c.border = { left: THIN, right: THIN, top: THIN, bottom: THIN }));
  });
  ps.getCell(rows.length + 4, 1).value = "※ 黄色のセル（単価・摘要）をご記入ください。金額・小計・合計は自動計算されます。";
  ps.getCell(rows.length + 4, 1).font = xf();
  return wb.xlsx.writeBuffer();
}

$("#export").addEventListener("click", async () => {
  const cats = catalog.filter(c => selected.has(c.id));
  if (!cats.length) return notify("出力する工事を選択してください。");
  const body = {
    property: propertyData(),
    categories: cats.map(c => ({
      name: c.name,
      items: c.items.filter(i => i.name.trim() && visible(i)).map(i => ({
        name: i.name, spec: i.spec || "", qty: i.group ? null : i.qty, unit: i.unit || "", note: i.note || "", group: !!i.group,
      })),
    })),
  };
  const pr = body.property, btn = $("#export");
  btn.disabled = true; btn.textContent = "作成中…";
  try {
    const buf = await buildXlsx(body);
    const filename = [pr.customer_name || "お客様", pr.work_title, "見積依頼書", pr.vendor_name].filter(Boolean).join("_") + ".xlsx";
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

function notify(msg) {
  const s = $("#status"); s.textContent = msg;
  setTimeout(() => renderDetail(), 3500);
}
