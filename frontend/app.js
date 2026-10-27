// 見積依頼書ジェネレーターの画面。
// 流れ: 物件情報を入力 → 工事(大項目)を選ぶ → 内訳を確認・編集 → Excel を出力
// web/build.py がこのファイルを単体HTML版に組み込む。そのとき「Excelを出力」から「起動」の手前までは
// web/export_xlsx.js に差し替わるので、両方で使う処理はそれより前に置く。

const API = (window.APP_CONFIG && window.APP_CONFIG.API_BASE || "").replace(/\/$/, "");
const SQM_PER_TSUBO = 3.30579;
const STRUCT = { wood: "木造", steel: "鉄骨造", rc: "RC造" };
const FLOOR_LABEL = { 1: "平屋", 2: "2階建て", 3: "3階建て" };
const $ = (s, r = document) => r.querySelector(s);
const esc = s => String(s ?? "").replace(/[&<>"']/g, m => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[m]));
const round2 = x => Math.round(x * 100) / 100;

// ---- 画面の状態 ----
let catalog = [];           // 大項目の一覧 [{id, name, items: [...]}]。items の形は backend/catalog.json を参照
let presets = [];           // 工事タイプ [{name, ids}]
const selected = new Set(); // 選択中の大項目 id
const openCats = new Set(); // 内訳で開いている大項目 id
const num = {};             // 数量の元になる数値(床面積など)。キーは入力欄の data-n
let structure = "wood";     // 構造(STRUCT のキー)
let floors = 2;             // 階数

// 構造・階数の条件に合う項目だけを表示・出力する
const visible = it => (!it.minFloors || it.minFloors <= floors) && (!it.structures || it.structures.includes(structure));
const visibleItems = c => c.items.filter(i => !i.group && visible(i));
const selectedCats = () => catalog.filter(c => selected.has(c.id));
const structLabel = () => `${STRUCT[structure]}・${FLOOR_LABEL[floors]}`;

// ---- 数量の自動入力 ----
// 入力された数値から、延床面積・坪数などの派生値も計算して返す
function basisValues() {
  const v = { ...num };
  if (floors < 3) v.floor3_area = null;
  if (floors < 2) v.floor2_area = null;
  const floorAreas = [v.floor1_area, v.floor2_area, v.floor3_area];
  if (floorAreas.some(x => x != null)) v.total_area = round2(floorAreas.reduce((s, x) => s + (x || 0), 0));
  if (v.total_area != null) v.tsubo = round2(v.total_area / SQM_PER_TSUBO);
  if (v.scaffold_area == null && v.wall_area != null) v.scaffold_area = v.wall_area; // 足場は空欄なら外壁面積
  return v;
}

// 数量 = 元の数値 × 係数。手で書き換えた数量(manual)はそのまま残す
function applyBasis() {
  const v = basisValues();
  for (const c of catalog)
    for (const it of c.items) {
      if (!it.basis || it.manual) continue;
      const b = v[it.basis];
      // 元の数値が空なら数量も空にする(消す前の値や、ひな形に残っている別案件の値を出さないため)
      it.qty = b == null ? null : round2(b * (it.factor || 1));
    }
}

// 元の数値が未入力で、数量を確認する必要がある項目か
const needsQty = (it, v) => it.basis && v[it.basis] == null && !it.manual;

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

// 工事タイプのボタン。今の選択と一致するタイプを強調する
function renderTypes() {
  const key = ids => [...ids].sort((a, b) => a - b).join();
  const cur = key(selected);
  $("#types").innerHTML = presets.map((p, i) =>
    `<button type="button" data-p="${i}" class="${key(p.ids) === cur ? "on" : ""}">${esc(p.name)}</button>`).join("");
}

// 工事(大項目)の選択カード
function renderCats() {
  renderTypes();
  $("#cats").innerHTML = catalog.map(c => `
    <label class="cat ${selected.has(c.id) ? "on" : ""}">
      <input type="checkbox" data-id="${c.id}" ${selected.has(c.id) ? "checked" : ""}>
      <span class="ico">${icon(c.id)}</span>
      ${esc(c.name)}<small>${visibleItems(c).length}項目</small>
    </label>`).join("");
}

// 内訳の表の1行。i は大項目の items 内の位置
function itemRow(it, i, v) {
  const cls = [it.group ? "group" : "", needsQty(it, v) ? "todo" : ""].join(" ");
  const cells = it.group ? `<td colspan="4"></td>` : `
    <td><input data-f="spec" value="${esc(it.spec)}"></td>
    <td class="q"><input data-f="qty" type="number" step="0.01" value="${esc(it.qty)}" title="${it.basis ? "物件情報から自動入力" : ""}"></td>
    <td class="u"><input data-f="unit" value="${esc(it.unit)}"></td>
    <td><input data-f="note" value="${esc(it.note)}"></td>`;
  return `
    <tr class="${cls}" data-i="${i}">
      <td><input data-f="name" value="${esc(it.name)}"></td>${cells}
      <td class="ac"><button class="mini" data-act="del" title="削除">✕</button><button class="mini" data-act="up" title="上へ">↑</button></td>
    </tr>`;
}

// 内訳(選択中の大項目ごとの表)
function renderDetail() {
  const v = basisValues();
  const cats = selectedCats();
  $("#detailBody").innerHTML = cats.length ? cats.map(c => `
    <details ${openCats.has(c.id) ? "open" : ""} data-cat="${c.id}">
      <summary><span class="no">${c.id}</span>${esc(c.name)}<small>${visibleItems(c).length}項目</small></summary>
      <div class="tbl"><table>
        <thead><tr><th>名称</th><th>仕様</th><th>数量</th><th>単位</th><th>摘要</th><th></th></tr></thead>
        <tbody>${c.items.map((it, i) => visible(it) ? itemRow(it, i, v) : "").join("")}</tbody>
      </table></div>
      <div class="tools"><button class="mini" data-act="add">＋ 項目を追加</button><button class="mini" data-act="addg">＋ 見出し(【 】)を追加</button></div>
    </details>`).join("") : `<p class="empty">「工事の選択」で工事を選ぶと、内訳がここに表示されます。</p>`;
  syncInputs(v);
  renderSummary(v, cats);
  const n = cats.reduce((s, c) => s + visibleItems(c).length, 0);
  $("#status").textContent = `${cats.length}工事 / ${n}項目 を出力`;
}

// 階数に合わせて床面積欄を、選択中の工事に必要な数値だけ入力欄を出す
function syncInputs(v) {
  const need = new Set();
  for (const c of selectedCats()) for (const it of c.items) if (it.basis && visible(it)) need.add(it.basis);
  document.querySelectorAll("[data-floor]").forEach(l => (l.hidden = +l.dataset.floor > floors));
  document.querySelectorAll("[data-need]").forEach(l => (l.hidden = !need.has(l.dataset.need)));
  $("#calc").textContent = v.total_area != null ? `延床面積 ${v.total_area}㎡（約 ${v.tsubo} 坪）` : "";
}

// 上部のタイトルとチップ
function renderTop() {
  const g = k => (document.querySelector(`[data-k=${k}]`)?.value || "").trim();
  const name = g("customer_name");
  $("#topTitle").textContent = name ? `${name} 様邸 ${g("work_title")}` : "新規の見積依頼";
  $("#topChips").innerHTML = [STRUCT[structure], FLOOR_LABEL[floors]].map(t => `<span class="chip">${t}</span>`).join("") +
    (g("vendor_name") ? `<span class="chip vendor">提出先 ${esc(g("vendor_name"))}</span>` : "");
}

// 右側の「出力の内容」と左ナビの補足
function renderSummary(v, cats) {
  const items = cats.flatMap(visibleItems);
  const todo = items.filter(i => needsQty(i, v)).length;
  $("#sStruct").textContent = structLabel();
  $("#sArea").textContent = v.total_area != null ? `${v.total_area}㎡（${v.tsubo}坪）` : "未入力";
  $("#sWorks").textContent = `${cats.length}工事`;
  $("#sItems").textContent = `${items.length}項目`;
  $("#sTodo").textContent = todo ? `${todo}項目` : "なし";
  $("#sTodo").className = todo ? "warn" : "";
  $("#navP").textContent = structLabel();
  $("#navW").textContent = `${cats.length}工事を選択`;
  $("#navD").textContent = `${items.length}項目`;
  renderTop();
}

function rerender() { renderCats(); renderDetail(); }

// ---- イベント: 物件情報 ----
$("#structures").addEventListener("change", e => {
  structure = e.target.value; applyBasis(); rerender();
});

$("#floors").addEventListener("change", e => {
  floors = +e.target.value; applyBasis(); rerender();
});

// data-k: 文字の項目(お客様名など) / data-n: 数量の元になる数値
document.addEventListener("input", e => {
  if (e.target.dataset.k) renderTop();
  const n = e.target.dataset.n;
  if (!n) return;
  num[n] = e.target.value === "" ? null : +e.target.value;
  applyBasis(); renderDetail();
});

// ---- イベント: 工事の選択 ----
$("#types").addEventListener("click", e => {
  const i = e.target.dataset.p; if (i == null) return;
  selected.clear(); presets[+i].ids.forEach(id => selected.add(id));
  rerender();
});

$("#cats").addEventListener("change", e => {
  const id = +e.target.dataset.id;
  e.target.checked ? selected.add(id) : selected.delete(id);
  rerender();
});

// ---- イベント: 内訳の編集 ----
const catOf = el => catalog.find(c => c.id === +el.closest("details").dataset.cat);

$("#detailBody").addEventListener("input", e => {
  const tr = e.target.closest("tr"); if (!tr) return;
  const it = catOf(e.target).items[+tr.dataset.i], f = e.target.dataset.f;
  if (f === "qty") { it.qty = e.target.value === "" ? null : +e.target.value; it.manual = true; }
  else it[f] = e.target.value;
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
    // 画面に出ている1つ上の項目と入れ替える(構造・階数で隠れている項目は飛ばす)
    let j = i - 1;
    while (j >= 0 && !visible(items[j])) j--;
    if (j >= 0) [items[j], items[i]] = [items[i], items[j]];
  }
  if (act === "add") items.push({ name: "", spec: "", qty: 1, unit: "式", manual: true });
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
  const v = basisValues();
  for (const k of ["total_area", "tsubo", "floor1_area", "floor2_area", "floor3_area", "roof_area", "wall_area", "scaffold_area", "toilet_count"])
    p[k] = v[k] ?? null;
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
    const res = await fetch(API + "/api/catalog", { method: "PUT", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ categories: clean, presets }) });
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
  catalog = cat.categories; presets = cat.presets || [];
  catalog.forEach(c => selected.add(c.id));
  applyBasis();
  rerender();
})();
