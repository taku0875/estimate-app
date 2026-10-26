const API = (window.APP_CONFIG && window.APP_CONFIG.API_BASE || "").replace(/\/$/, "");
const SQM_PER_TSUBO = 3.30579;
const $ = (s, r = document) => r.querySelector(s);

let catalog = [];          // [{id,name,items:[...]}]
const selected = new Set(); // 選択中の大項目id
const num = {};             // 数量の元になる数値
let structure = "wood";   // 構造
let presets = [];
let floors = 2;             // 階数
const openCats = new Set(); // 開いている大項目
const STRUCT = { wood: "木造", steel: "鉄骨造", rc: "RC造" };
const visible = it => (!it.minFloors || it.minFloors <= floors) && (!it.structures || it.structures.includes(structure));

// ---- 数量の元になる数値 ----
function basisValues() {
  const v = { ...num };
  const fl = [v.floor1_area, v.floor2_area, v.floor3_area].slice(0, floors);
  if (fl.some(x => x != null)) v.total_area = Math.round(fl.reduce((s, x) => s + (x || 0), 0) * 100) / 100;
  if (v.floor2_area != null && floors < 2) v.floor2_area = null;
  if (v.floor3_area != null && floors < 3) v.floor3_area = null;
  if (v.total_area != null) v.tsubo = Math.round((v.total_area / SQM_PER_TSUBO) * 100) / 100;
  if (v.scaffold_area == null && v.wall_area != null) v.scaffold_area = v.wall_area;
  return v;
}

function applyBasis() {
  const v = basisValues();
  for (const c of catalog)
    for (const it of c.items) {
      if (!it.basis || it.manual) continue;
      const b = v[it.basis];
      if (b != null) it.qty = Math.round(b * (it.factor || 1) * 100) / 100;
    }
}

// ---- 描画 ----
function renderTypes() {
  const cur = [...selected].sort((a, b) => a - b).join();
  $("#types").innerHTML = presets.map((p, i) =>
    `<button type="button" data-p="${i}" class="${[...p.ids].sort((a, b) => a - b).join() === cur ? "on" : ""}">${esc(p.name)}</button>`).join("");
}

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

function renderCats() {
  renderTypes();
  $("#cats").innerHTML = catalog.map(c => `
    <label class="cat ${selected.has(c.id) ? "on" : ""}">
      <input type="checkbox" data-id="${c.id}" ${selected.has(c.id) ? "checked" : ""}>
      <span class="ico">${icon(c.id)}</span>
      ${esc(c.name)}<small>${c.items.filter(i => !i.group && visible(i)).length}項目</small>
    </label>`).join("");
}

function renderDetail() {
  const v = basisValues();
  const cats = catalog.filter(c => selected.has(c.id));
  $("#detailBody").innerHTML = cats.length ? cats.map(c => `
    <details ${openCats.has(c.id) ? "open" : ""} data-cat="${c.id}">
      <summary><span class="no">${c.id}</span>${esc(c.name)}<small>${c.items.filter(i => !i.group && visible(i)).length}項目</small></summary>
      <div class="tbl"><table>
        <thead><tr><th>名称</th><th>仕様</th><th>数量</th><th>単位</th><th>摘要</th><th></th></tr></thead>
        <tbody>
        ${c.items.map((it, i) => !visible(it) ? "" : `
          <tr class="${it.group ? "group" : ""} ${it.basis && v[it.basis] == null && !it.manual ? "todo" : ""}" data-i="${i}">
            <td><input data-f="name" value="${esc(it.name)}"></td>
            ${it.group ? `<td colspan="4"></td>` : `
            <td><input data-f="spec" value="${esc(it.spec || "")}"></td>
            <td class="q"><input data-f="qty" type="number" step="0.01" value="${esc(it.qty ?? "")}" title="${it.basis ? "物件情報から自動入力" : ""}"></td>
            <td class="u"><input data-f="unit" value="${esc(it.unit || "")}"></td>
            <td><input data-f="note" value="${esc(it.note || "")}"></td>`}
            <td class="ac"><button class="mini" data-act="del" title="削除">✕</button><button class="mini" data-act="up" title="上へ">↑</button></td>
          </tr>`).join("")}
        </tbody>
      </table></div>
      <div class="tools"><button class="mini" data-act="add">＋ 項目を追加</button><button class="mini" data-act="addg">＋ 見出し(【 】)を追加</button></div>
    </details>`).join("") : `<p class="empty">「工事の選択」で工事を選ぶと、内訳がここに表示されます。</p>`;
  const n = cats.reduce((s, c) => s + c.items.filter(i => !i.group && visible(i)).length, 0);
  syncInputs(v);
  renderSummary(v, cats);
  $("#status").textContent = `${cats.length}工事 / ${n}項目 を出力`;
}

function syncInputs(v) {
  const need = new Set();
  for (const c of catalog) if (selected.has(c.id)) for (const it of c.items) if (it.basis && visible(it)) need.add(it.basis);
  document.querySelectorAll("[data-floor]").forEach(l => (l.hidden = +l.dataset.floor > floors));
  document.querySelectorAll("[data-need]").forEach(l => (l.hidden = !need.has(l.dataset.need)));
  $("#calc").textContent = v.total_area != null ? `延床面積 ${v.total_area}㎡（約 ${v.tsubo} 坪）` : "";
}

function renderTop() {
  const g = k => (document.querySelector(`[data-k=${k}]`)?.value || "").trim();
  const name = g("customer_name");
  $("#topTitle").textContent = name ? `${name} 様邸 ${g("work_title")}` : "新規の見積依頼";
  const chips = [STRUCT[structure], ["", "平屋", "2階建て", "3階建て"][floors]];
  $("#topChips").innerHTML = chips.map(t => `<span class="chip">${t}</span>`).join("") +
    (g("vendor_name") ? `<span class="chip vendor">提出先 ${esc(g("vendor_name"))}</span>` : "");
}

function renderSummary(v, cats) {
  const items = cats.flatMap(c => c.items.filter(i => !i.group && visible(i)));
  const todo = items.filter(i => i.basis && v[i.basis] == null && !i.manual).length;
  $("#sStruct").textContent = `${STRUCT[structure]}・${["", "平屋", "2階建て", "3階建て"][floors]}`;
  $("#sArea").textContent = v.total_area != null ? `${v.total_area}㎡（${v.tsubo}坪）` : "未入力";
  $("#sWorks").textContent = `${cats.length}工事`;
  $("#sItems").textContent = `${items.length}項目`;
  $("#sTodo").textContent = todo ? `${todo}項目` : "なし";
  $("#sTodo").className = todo ? "warn" : "";
  $("#navP").textContent = `${STRUCT[structure]}・${["", "平屋", "2階建て", "3階建て"][floors]}`;
  $("#navW").textContent = `${cats.length}工事を選択`;
  $("#navD").textContent = `${items.length}項目`;
  renderTop();
}

const esc = s => String(s ?? "").replace(/[&<>"']/g, m => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[m]));

// ---- イベント ----
$("#cats").addEventListener("change", e => {
  const id = +e.target.dataset.id;
  e.target.checked ? selected.add(id) : selected.delete(id);
  renderCats(); renderDetail();
});

$("#detailBody").addEventListener("input", e => {
  const tr = e.target.closest("tr"); if (!tr) return;
  const cat = catalog.find(c => c.id === +e.target.closest("details").dataset.cat);
  const it = cat.items[+tr.dataset.i], f = e.target.dataset.f;
  if (f === "qty") { it.qty = e.target.value === "" ? null : +e.target.value; it.manual = true; }
  else it[f] = e.target.value;
});

$("#detailBody").addEventListener("toggle", e => {
  const id = +e.target.dataset.cat;
  e.target.open ? openCats.add(id) : openCats.delete(id);
}, true);

$("#types").addEventListener("click", e => {
  const i = e.target.dataset.p; if (i == null) return;
  selected.clear(); presets[+i].ids.forEach(id => selected.add(id));
  renderCats(); renderDetail();
});

$("#structures").addEventListener("change", e => {
  structure = e.target.value; applyBasis(); renderCats(); renderDetail();
});

$("#floors").addEventListener("change", e => {
  floors = +e.target.value; applyBasis(); renderCats(); renderDetail();
});

$("#detailBody").addEventListener("click", e => {
  const act = e.target.dataset.act; if (!act) return;
  const cat = catalog.find(c => c.id === +e.target.closest("details").dataset.cat);
  const tr = e.target.closest("tr"), i = tr ? +tr.dataset.i : -1;
  if (act === "del") cat.items.splice(i, 1);
  if (act === "up" && i > 0) [cat.items[i - 1], cat.items[i]] = [cat.items[i], cat.items[i - 1]];
  if (act === "add") cat.items.push({ name: "", spec: "", qty: 1, unit: "式", manual: true });
  if (act === "addg") cat.items.push({ name: "【】", group: true });
  renderDetail();
});

document.addEventListener("input", e => {
  if (e.target.dataset.k) renderTop();
  const n = e.target.dataset.n;
  if (!n) return;
  num[n] = e.target.value === "" ? null : +e.target.value;
  applyBasis(); renderDetail();
});

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

$("#export").addEventListener("click", async () => {
  const cats = catalog.filter(c => selected.has(c.id));
  if (!cats.length) return alert("出力する工事を選択してください。");
  const body = {
    property: propertyData(),
    categories: cats.map(c => ({
      name: c.name,
      items: c.items.filter(i => i.name.trim() && visible(i)).map(i => ({
        name: i.name, spec: i.spec || "", qty: i.group ? null : i.qty, unit: i.unit || "", note: i.note || "", group: !!i.group,
      })),
    })),
  };
  const btn = $("#export"); btn.disabled = true; btn.textContent = "作成中…";
  try {
    const res = await fetch(API + "/api/export", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
    if (!res.ok) throw new Error(await res.text());
    const blob = await res.blob();
    const name = decodeURIComponent((res.headers.get("Content-Disposition") || "").split("''")[1] || "見積依頼書.xlsx");
    const a = Object.assign(document.createElement("a"), { href: URL.createObjectURL(blob), download: name });
    a.click(); URL.revokeObjectURL(a.href);
  } catch (err) { alert("出力に失敗しました: " + err.message); }
  finally { btn.disabled = false; btn.textContent = "Excelを出力"; }
});

$("#saveCatalog").addEventListener("click", async () => {
  if (!confirm("現在の項目構成をひな形として保存します。よろしいですか？")) return;
  const clean = catalog.map(c => ({ ...c, items: c.items.map(({ manual, ...r }) => r) }));
  try {
    const res = await fetch(API + "/api/catalog", { method: "PUT", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ categories: clean, presets }) });
    alert(res.ok ? "保存しました。" : res.status === 403 ? "このサーバーではひな形の保存が無効になっています。" : "保存に失敗しました。");
  } catch (err) { alert("保存に失敗しました: " + err.message); }
});

// ---- 左ナビ: 表示中のセクションを強調 ----
const navLinks = [...document.querySelectorAll("#nav a")];
["property", "works", "detail"].forEach(id => {
  new IntersectionObserver(es => es.forEach(en => en.isIntersecting &&
    navLinks.forEach(a => a.classList.toggle("on", a.getAttribute("href") === "#" + id))),
  { rootMargin: "-20% 0px -65% 0px" }).observe(document.getElementById(id));
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
  renderCats(); renderDetail();
})();
