// TSUMIKIの画面。
// 流れ: 物件情報を入力 → 工事(大項目)と水回りの範囲を選ぶ → 内訳に数量を入力 → Excel を出力
// web/build.py がこのファイルを単体HTML版に組み込む。そのとき「Excelを出力」から「起動」の手前までは
// web/export_xlsx.js に差し替わるので、両方で使う処理はそれより前に置く。

const API = (window.APP_CONFIG && window.APP_CONFIG.API_BASE || "").replace(/\/$/, "");
const STRUCT = { wood: "木造", steel: "鉄骨造", rc: "RC造" };
const FLOOR_LABEL = { 1: "平屋", 2: "2階建て", 3: "3階建て" };
// Excelの形式(見本のExcelに合わせた2種類)。工事名の初期値も形式ごとに持つ
const FORMATS = {
  renovation: { name: "リノベーション工事", title: "リノベーション工事" },
  mizumawari: { name: "水回りリフォーム", title: "水回りリフォーム工事" },
  exterior: { name: "外装工事", title: "外装工事" },
  interior: { name: "内装工事", title: "内装工事" },
};
const $ = (s, r = document) => r.querySelector(s);
const esc = s => String(s ?? "").replace(/[&<>"']/g, m => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[m]));

// ---- 画面の状態 ----
let catalog = [];           // 大項目の一覧 [{id, name, items: [...]}]。items の形は backend/catalog.json を参照
let presets = [];           // 工事の種類ごとの大項目 [{name, ids, format}]
let plumbing = { fixtures: [] }; // 水回りの設備(キッチン等)
const selected = new Set(); // 選択中の大項目 id
const openCats = new Set(); // 内訳で開いている大項目 id
const fixturesOn = new Set(); // 水回りで選択中の設備のキー(kitchen など)
let structure = "wood";     // 構造(STRUCT のキー)
let floors = 2;             // 階数
let format = "renovation";  // Excelの形式(FORMATS のキー)

// 構造・階数・水回りの範囲に合う項目だけを表示・出力する。fixtures を持つ項目は、そのどれかが選ばれていれば出す
const visible = it =>
  (!it.minFloors || it.minFloors <= floors) &&
  (!it.structures || it.structures.includes(structure)) &&
  (!it.fixtures || it.fixtures.some(f => fixturesOn.has(f)));
const visibleItems = c => c.items.filter(i => !i.group && visible(i));
const selectedCats = () => catalog.filter(c => selected.has(c.id));
const structLabel = () => `${STRUCT[structure]}・${FLOOR_LABEL[floors]}`;

// 数量は案件ごとに手で入力する。未入力(空)の項目は内訳で黄色にして、サマリに件数を出す
const needsQty = it => !it.group && (it.qty == null || it.qty === "");

// 「水回りリフォーム」で、選択中の工事に水回りの項目が含まれるときだけ、設備の選択を出す。他の種類では設備はすべて対象
const hasFixtureItems = () => format === "mizumawari" && selectedCats().some(c => c.items.some(i => i.fixtures && (!i.minFloors || i.minFloors <= floors)));
const plumbingLabel = () => {
  if (!hasFixtureItems()) return "";
  const names = plumbing.fixtures.filter(f => fixturesOn.has(f.key)).map(f => f.name);
  return names.length ? names.join("＋") : "なし";
};

// ---- 描画 ----
const ICONS = {
  1: '<path d="M14 4l6 6-3 3-6-6zM12 8L4 16l4 4 8-8"/>',
  2: '<path d="M3 8l9-4 9 4-9 4zM3 13l9 4 9-4M3 17l9 4 9-4"/>',
  3: '<path d="M4 20V8l8-4 8 4v12M9 20v-6h6v6M4 12h16"/>',
  4: '<path d="M2 12l10-8 10 8M5 10v10h14V10"/>',
  5: '<rect x="4" y="4" width="16" height="16" rx="1"/><path d="M4 12h16M12 4v8M8 12v8M16 12v8"/>',
  6: '<rect x="5" y="3" width="14" height="18" rx="1"/><path d="M12 3v18M5 12h14"/>',
  7: '<path d="M4 4h12v6H4zM16 7h4v6h-8v3M10 16h4v5h-4z"/>',
  8: '<path d="M13 2L5 14h6l-1 8 8-12h-6z"/>',
  9: '<path d="M12 3s6 6.5 6 11a6 6 0 0 1-12 0c0-4.5 6-11 6-11z"/>',
  10: '<path d="M3 12h18v3a5 5 0 0 1-5 5H8a5 5 0 0 1-5-5zM6 12V6a2 2 0 0 1 4 0M5 20l-1 2M19 20l1 2"/>',
  11: '<path d="M5 3v18M19 3v18M5 7h14M5 12h14M5 17h14"/>',
  12: '<path d="M6 3h12v18l-3-2-3 2-3-2-3 2zM9 8h6M9 12h6"/>',
};
const icon = id => `<svg viewBox="0 0 24 24" aria-hidden="true">${ICONS[id] || ICONS[12]}</svg>`;

const sameSet = (a, b) => a.size === b.size && [...a].every(x => b.has(x));

// 水回りの設備(水回りリフォームのときだけ表示)のチェックボックス
function renderPlumbing() {
  const box = $("#plumbing");
  box.hidden = !hasFixtureItems();
  if (box.hidden) return;
  $("#fixtures").innerHTML = plumbing.fixtures.map(f =>
    `<label><input type="checkbox" data-key="${f.key}" ${fixturesOn.has(f.key) ? "checked" : ""}>${esc(f.name)}</label>`).join("");
}

// 工事の種類(=Excelの形式)のボタン
function renderFormats() {
  $("#formats").innerHTML = Object.entries(FORMATS).map(([k, f]) =>
    `<button type="button" data-f="${k}" class="${k === format ? "on" : ""}">${f.name}</button>`).join("");
}

// 種類を切り替える。その種類の大項目を選び、工事名が初期値のままなら種類に合う初期値に直す
function setFormat(k) {
  const t = $("[data-k=work_title]");
  if (t && Object.values(FORMATS).some(f => f.title === t.value.trim())) t.value = FORMATS[k].title;
  format = k;
  plumbing.fixtures.forEach(f => fixturesOn.add(f.key)); // 設備は、種類を変えるたびにすべて対象に戻す
  const p = presets.find(x => x.format === k);
  if (p) { selected.clear(); p.ids.forEach(id => selected.add(id)); }
  rerender();
}

// 工事(大項目)の選択カード
function renderCats() {
  renderFormats();
  renderPlumbing();
  $("#cats").innerHTML = catalog.map(c => `
    <label class="cat ${selected.has(c.id) ? "on" : ""}">
      <input type="checkbox" data-id="${c.id}" ${selected.has(c.id) ? "checked" : ""}>
      <span class="ico">${icon(c.id)}</span>
      ${esc(c.name)}<small>${visibleItems(c).length}項目</small>
    </label>`).join("");
}

// 内訳の表の1行。i は大項目の items 内の位置
function itemRow(it, i) {
  const cls = [it.group ? "group" : "", needsQty(it) ? "todo" : ""].join(" ");
  const cells = it.group ? `<td colspan="4"></td>` : `
    <td><input data-f="spec" value="${esc(it.spec)}" placeholder="仕様を記入"></td>
    <td class="q"><input data-f="qty" type="number" step="any" inputmode="decimal" placeholder="数量" value="${esc(it.qty)}"></td>
    <td class="u"><input data-f="unit" value="${esc(it.unit)}"></td>
    <td><input data-f="note" value="${esc(it.note)}" placeholder="摘要"></td>`;
  return `
    <tr class="${cls}" data-i="${i}">
      <td><input data-f="name" value="${esc(it.name)}"></td>${cells}
      <td class="ac"><button class="mini" data-act="del" title="削除">✕</button><button class="mini" data-act="up" title="上へ">↑</button><button class="mini" data-act="down" title="下へ">↓</button></td>
    </tr>`;
}

// 内訳(選択中の大項目ごとの表)
function renderDetail() {
  const cats = selectedCats();
  $("#detailBody").innerHTML = cats.length ? cats.map(c => `
    <details ${openCats.has(c.id) ? "open" : ""} data-cat="${c.id}">
      <summary><span class="no">${c.id}</span>${esc(c.name)}<small>${visibleItems(c).length}項目</small></summary>
      <div class="tbl"><table>
        <thead><tr><th>名称</th><th>仕様</th><th>数量</th><th>単位</th><th>摘要</th><th></th></tr></thead>
        <tbody>${c.items.map((it, i) => visible(it) ? itemRow(it, i) : "").join("")}</tbody>
      </table></div>
      <div class="tools"><button class="mini" data-act="add">＋ 項目を追加</button><button class="mini" data-act="addg">＋ 見出し(【 】)を追加</button></div>
    </details>`).join("") : `<p class="empty">「工事の選択」で工事を選ぶと、内訳がここに表示されます。</p>`;
  renderSummary(cats);
  const n = cats.reduce((s, c) => s + visibleItems(c).length, 0);
  $("#status").textContent = `${cats.length}工事 / ${n}項目 を出力`;
}

// 上部のタイトルとチップ
function renderTop() {
  const g = k => (document.querySelector(`[data-k=${k}]`)?.value || "").trim();
  const name = g("customer_name");
  $("#topTitle").textContent = name ? `${name} 様邸 ${g("work_title")}` : "新規の見積依頼";
  $("#topChips").innerHTML = [STRUCT[structure], FLOOR_LABEL[floors]].map(t => `<span class="chip">${t}</span>`).join("") +
    "";
}

// 右側の「出力の内容」と左ナビの補足
function renderSummary(cats) {
  const items = cats.flatMap(visibleItems);
  const todo = items.filter(needsQty).length;
  const pl = plumbingLabel();
  $("#sStruct").textContent = structLabel();
  $("#sPlumb").textContent = pl || "—";
  $("#sWorks").textContent = `${cats.length}工事`;
  $("#sItems").textContent = `${items.length}項目`;
  $("#sTodo").textContent = todo ? `${todo}項目` : "なし";
  $("#sTodo").className = todo ? "warn" : "ok";
  $("#navP").textContent = structLabel();
  $("#navW").textContent = `${cats.length}工事を選択`;
  $("#navD").textContent = `${items.length}項目`;
  renderTop();
}

function rerender() { renderCats(); renderDetail(); }

// ---- イベント: 物件情報 ----
$("#structures").addEventListener("change", e => { structure = e.target.value; rerender(); });
$("#floors").addEventListener("change", e => { floors = +e.target.value; rerender(); });
document.addEventListener("input", e => { if (e.target.dataset.k) renderTop(); }); // data-k: 文字の項目(お客様名など)

// ---- イベント: 工事の選択 ----
$("#formats").addEventListener("click", e => { const k = e.target.dataset.f; if (k) setFormat(k); });

$("#cats").addEventListener("change", e => {
  const id = +e.target.dataset.id;
  e.target.checked ? selected.add(id) : selected.delete(id);
  rerender();
});

// 水回り: 設備のチェックで、給排水設備工事・住宅設備工事などの内訳が変わる
$("#fixtures").addEventListener("change", e => {
  const k = e.target.dataset.key;
  e.target.checked ? fixturesOn.add(k) : fixturesOn.delete(k);
  rerender();
});

// ---- イベント: 内訳の編集 ----
const catOf = el => catalog.find(c => c.id === +el.closest("details").dataset.cat);

$("#detailBody").addEventListener("input", e => {
  const tr = e.target.closest("tr"); if (!tr) return;
  const it = catOf(e.target).items[+tr.dataset.i], f = e.target.dataset.f;
  if (f === "qty") {
    it.qty = e.target.value === "" ? null : +e.target.value;
    tr.classList.toggle("todo", needsQty(it)); // 入力のたびに表を作り直すとカーソルが外れるので、行の印と件数だけ更新する
    renderSummary(selectedCats());
  } else it[f] = e.target.value;
});

$("#detailBody").addEventListener("toggle", e => {
  const id = +e.target.dataset.cat;
  e.target.open ? openCats.add(id) : openCats.delete(id);
}, true);

$("#detailBody").addEventListener("click", e => {
  const act = e.target.dataset.act; if (!act) return;
  const items = catOf(e.target).items;
  const i = +(e.target.closest("tr")?.dataset.i ?? -1);
  if (act === "del") items.splice(i, 1);
  if (act === "up") {
    // 画面に出ている1つ上の項目と入れ替える(構造・階数などで隠れている項目は飛ばす)
    let j = i - 1;
    while (j >= 0 && !visible(items[j])) j--;
    if (j >= 0) [items[j], items[i]] = [items[i], items[j]];
  }
  if (act === "down") {
    // 画面に出ている1つ下の項目と入れ替える
    let j = i + 1;
    while (j < items.length && !visible(items[j])) j++;
    if (j < items.length) [items[j], items[i]] = [items[i], items[j]];
  }
  if (act === "add") items.push({ name: "", spec: "", qty: null, unit: "式" });
  if (act === "addg") items.push({ name: "【】", group: true });
  renderDetail();
});

// ---- 左ナビ: 表示中のセクションを強調 ----
const navLinks = [...document.querySelectorAll("#nav a")];
["property", "works", "detail"].forEach(id => {
  new IntersectionObserver(es => es.forEach(en => en.isIntersecting &&
    navLinks.forEach(a => a.classList.toggle("on", a.getAttribute("href") === "#" + id))),
  { rootMargin: "-20% 0px -65% 0px" }).observe(document.getElementById(id));
});

// ---- 出力するデータ(バックエンドの ExportRequest と同じ形) ----
function propertyData() {
  const p = {};
  document.querySelectorAll("[data-k]").forEach(el => (p[el.dataset.k] = el.value));
  p.floors = floors;
  p.structure = STRUCT[structure];
  p.plumbing = plumbingLabel();
  p.format = format;
  return p;
}

function exportData() {
  return {
    property: propertyData(),
    categories: selectedCats().map(c => ({
      name: c.name,
      items: c.items.filter(i => i.name.trim() && visible(i)).map(i => ({
        name: i.name, spec: i.spec || "", qty: i.group ? null : i.qty, unit: i.unit || "", note: i.note || "", group: !!i.group,
      })),
    })),
  };
}

// ---- Excelを出力・ひな形として保存(サーバー版) ----
$("#export").addEventListener("click", async () => {
  if (!selected.size) return alert("出力する工事を選択してください。");
  const btn = $("#export"); btn.disabled = true; btn.textContent = "作成中…";
  try {
    const res = await fetch(API + "/api/export", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(exportData()) });
    if (!res.ok) throw new Error(res.status === 422 ? "入力が長すぎるか、数値が正しくありません。" : `サーバーエラー(${res.status})`);
    const blob = await res.blob();
    const name = decodeURIComponent((res.headers.get("Content-Disposition") || "").split("''")[1] || "見積依頼書.xlsx");
    const a = Object.assign(document.createElement("a"), { href: URL.createObjectURL(blob), download: name });
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  } catch (err) { alert("出力に失敗しました: " + err.message); }
  finally { btn.disabled = false; btn.textContent = "Excelを出力"; }
});

$("#saveCatalog").addEventListener("click", async () => {
  if (!confirm("現在の項目構成をひな形として保存します。よろしいですか？")) return;
  const clean = catalog.map(c => ({ ...c, items: c.items.map(({ manual, ...r }) => r) })); // 画面用の印(manual)は保存しない
  try {
    const res = await fetch(API + "/api/catalog", { method: "PUT", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ categories: clean, presets, plumbing }) });
    alert(res.ok ? "保存しました。" : res.status === 403 ? "このサーバーではひな形の保存が無効になっています。" : "保存に失敗しました。");
  } catch (err) { alert("保存に失敗しました: " + err.message); }
});

// ---- 起動 ----
(async () => {
  let cat;
  try {
    const res = await fetch(API + "/api/catalog");
    if (!res.ok) throw new Error(res.status);
    cat = await res.json();
  } catch (err) {
    $("#detailBody").innerHTML = `<p class="empty">工事項目を読み込めませんでした。サーバーの起動と config.js の API_BASE を確認してください。</p>`;
    return;
  }
  catalog = cat.categories; presets = cat.presets || []; plumbing = cat.plumbing || plumbing;
  // ひな形に残っている別案件の数量は使わない。数量は案件ごとに手で入力する(「式」だけは 1)
  for (const c of catalog) for (const it of c.items) if (!it.group) it.qty = it.unit === "式" ? 1 : null;
  catalog.forEach(c => selected.add(c.id)); // 初期は「リノベーション工事」(すべての大項目)
  plumbing.fixtures.forEach(f => fixturesOn.add(f.key)); // 初期は設備すべて
  rerender();
})();
