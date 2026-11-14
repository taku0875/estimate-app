"use client";
// TSUMIKIの画面(Next.js 版)。frontend/app.js と同じ流れ:
// 物件情報 → 工事と水回りの範囲 → 内訳に数量を入力 → Excel を出力(FastAPI の /api/export)
import { useEffect, useMemo, useState } from "react";
import type { Category, CatalogData, ExcelFormat, Floors, Item, Plumbing, Preset, Prop, Structure } from "@/lib/types";
import {
  API, FLOOR_LABEL, FORMATS, ICONS, STRUCT, hasFixtureItems, needsQty, plumbingLabel, sameSet, visible, visibleItems, type Cond,
} from "@/lib/domain";

const COMPANY = "株式会社石丸ハウスセンター"; // 依頼元(自社名)は画面では入力しない
const EMPTY: Prop = {
  customer_company: "", customer_name: "", address: "", start_date: "", end_date: "",
  work_period: "", due_date: "", remarks: "",
};

export default function Estimator() {
  const [catalog, setCatalog] = useState<Category[]>([]);
  const [presets, setPresets] = useState<Preset[]>([]);
  const [plumbing, setPlumbing] = useState<Plumbing>({ fixtures: [] });
  const [selected, setSelected] = useState<Set<number>>(new Set());
  const [openCats, setOpenCats] = useState<Set<number>>(new Set());
  const [fixturesOn, setFixturesOn] = useState<Set<string>>(new Set());
  const [structure, setStructure] = useState<Structure>("wood");
  const [floors, setFloors] = useState<Floors>(2);
  const [format, setFormat] = useState<ExcelFormat>("renovation");
  const [prop, setProp] = useState<Prop>(EMPTY);
  const [loadError, setLoadError] = useState(false);
  const [busy, setBusy] = useState(false);
  const [msg, setMsg] = useState("");
  const [active, setActive] = useState("property");

  // ---- 起動: ひな形(catalog)を API から読む ----
  useEffect(() => {
    (async () => {
      try {
        const res = await fetch(API + "/api/catalog");
        if (!res.ok) throw new Error(String(res.status));
        const cat: CatalogData = await res.json();
        // 別案件の数量は使わない。数量は案件ごとに手で入力する(「式」だけは 1)
        const cats = cat.categories.map(c => ({
          ...c, items: c.items.map(it => (it.group ? it : { ...it, qty: it.unit === "式" ? 1 : null })),
        }));
        const pl = cat.plumbing ?? { fixtures: [] };
        setCatalog(cats);
        setPresets(cat.presets ?? []);
        setPlumbing(pl);
        setSelected(new Set(cats.map(c => c.id)));
        setFixturesOn(new Set(pl.fixtures.map(f => f.key))); // 初期は設備すべて
      } catch {
        setLoadError(true);
      }
    })();
  }, []);

  // ---- 左ナビ: 表示中のセクションを強調 ----
  useEffect(() => {
    const obs = ["property", "works", "detail"].map(id => {
      const el = document.getElementById(id);
      if (!el) return null;
      const io = new IntersectionObserver(es => es.forEach(e => e.isIntersecting && setActive(id)),
        { rootMargin: "-20% 0px -65% 0px" });
      io.observe(el);
      return io;
    });
    return () => obs.forEach(o => o?.disconnect());
  }, []);

  const cond: Cond = useMemo(() => ({ structure, floors, fixturesOn }), [structure, floors, fixturesOn]);
  const cats = catalog.filter(c => selected.has(c.id));
  const items = cats.flatMap(c => visibleItems(c, cond));
  const todo = items.filter(needsQty).length;
  const showPlumb = hasFixtureItems(cats, floors, format);
  const plLabel = plumbingLabel(cats, plumbing, cond, format);
  const structLabel = `${STRUCT[structure]}・${FLOOR_LABEL[floors]}`;

  const setP = (k: keyof Prop) => (e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>) =>
    setProp(p => ({ ...p, [k]: e.target.value }));
  const toggle = <T,>(s: Set<T>, v: T, on: boolean) => { const n = new Set(s); on ? n.add(v) : n.delete(v); return n; };

  // 内訳の編集。ci: 大項目の id、i: items 内の位置
  const patchCat = (ci: number, fn: (items: Item[]) => Item[]) =>
    setCatalog(cs => cs.map(c => (c.id === ci ? { ...c, items: fn(c.items) } : c)));
  const patchItem = (ci: number, i: number, patch: Partial<Item>) =>
    patchCat(ci, its => its.map((it, k) => (k === i ? { ...it, ...patch } : it)));
  const moveDown = (ci: number, i: number) => patchCat(ci, its => {
    // 画面に出ている1つ下の項目と入れ替える
    let j = i + 1;
    while (j < its.length && !visible(its[j], cond)) j++;
    if (j >= its.length) return its;
    const n = [...its]; [n[j], n[i]] = [n[i], n[j]]; return n;
  });
  const moveUp = (ci: number, i: number) => patchCat(ci, its => {
    // 画面に出ている1つ上の項目と入れ替える(隠れている項目は飛ばす)
    let j = i - 1;
    while (j >= 0 && !visible(its[j], cond)) j--;
    if (j < 0) return its;
    const n = [...its]; [n[j], n[i]] = [n[i], n[j]]; return n;
  });

  // 種類を切り替える。その種類の大項目を選ぶ
  const changeFormat = (k: ExcelFormat) => {
    const p = presets.find(x => x.format === k);
    if (p) setSelected(new Set(p.ids));
    setFixturesOn(new Set(plumbing.fixtures.map(f => f.key))); // 設備は、種類を変えるたびにすべて対象に戻す
    setFormat(k);
  };

  const flash = (m: string) => { setMsg(m); setTimeout(() => setMsg(""), 3500); };

  // ---- 出力するデータ(バックエンドの ExportRequest と同じ形) ----
  const exportData = () => ({
    property: { ...prop, company_name: COMPANY, floors, structure: STRUCT[structure], plumbing: plLabel, format },
    categories: cats.map(c => ({
      name: c.name,
      items: c.items.filter(i => i.name.trim() && visible(i, cond)).map(i => ({
        name: i.name, spec: i.spec || "", qty: i.group ? null : i.qty ?? null, unit: i.unit || "", note: i.note || "", group: !!i.group,
      })),
    })),
  });

  async function doExport() {
    if (!selected.size) return alert("出力する工事を選択してください。");
    setBusy(true);
    try {
      const res = await fetch(API + "/api/export", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(exportData()) });
      if (!res.ok) throw new Error(res.status === 422 ? "入力が長すぎるか、数値が正しくありません。" : `サーバーエラー(${res.status})`);
      const blob = await res.blob();
      const name = decodeURIComponent((res.headers.get("Content-Disposition") || "").split("''")[1] || "見積依頼書.xlsx");
      const a = Object.assign(document.createElement("a"), { href: URL.createObjectURL(blob), download: name });
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(a.href), 4000);
      flash("Excelを出力しました。");
    } catch (err) {
      alert("出力に失敗しました: " + (err instanceof Error ? err.message : err));
    } finally { setBusy(false); }
  }

  async function saveCatalog() {
    if (!confirm("現在の項目構成をひな形として保存します。よろしいですか？")) return;
    try {
      const res = await fetch(API + "/api/catalog", { method: "PUT", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ categories: catalog, presets, plumbing }) });
      alert(res.ok ? "保存しました。" : res.status === 403 ? "このサーバーではひな形の保存が無効になっています。" : "保存に失敗しました。");
    } catch (err) { alert("保存に失敗しました: " + (err instanceof Error ? err.message : err)); }
  }

  const field = (k: keyof Prop, label: string, extra: React.InputHTMLAttributes<HTMLInputElement> = {}, cls = "") => (
    <label className={cls}>{label}<input value={prop[k]} onChange={setP(k)} {...extra} /></label>
  );

  return (
    <div className="app">
      <aside className="side">
        <div className="brand">
          <svg className="logo" viewBox="0 0 32 32" aria-hidden="true"><rect width="32" height="32" rx="9" fill="#3b4533" /><rect x="5" y="17" width="10" height="10" rx="2" fill="#a9d46a" /><rect x="17" y="17" width="10" height="10" rx="2" fill="#a9d46a" /><rect x="11" y="5" width="10" height="10" rx="2" fill="#a9d46a" /></svg>
          <div><b>TSUMIKI</b></div>
        </div>
        <nav id="nav">
          {[["property", "1", "物件情報"], ["works", "2", "工事の選択"], ["detail", "3", "内訳の確認"]].map(([id, n, t]) => (
            <a key={id} href={`#${id}`} className={active === id ? "on" : ""}><i>{n}</i><span>{t}</span></a>
          ))}
        </nav>
      </aside>

      <div className="content">
        <header className="top">
          <div className="crumb"><b>新規の見積依頼</b></div>
        </header>

        <main>
          <section className="card" id="property">
            <div className="card-h"><h2>物件情報</h2></div>
            <div className="grid">
              <div className="wide pair">
                <span className="gl">お施主様名</span>
                <div className="two">
                  {field("customer_company", "法人の名称", { placeholder: "例: 株式会社〇〇" })}
                  {field("customer_name", "個人の名前", { placeholder: "例: 山田太郎" })}
                </div>
              </div>
              {field("address", "物件所在地", { placeholder: "例: 名古屋市中区…" }, "wide")}
              {field("start_date", "着工予定日", { type: "date" })}
              {field("end_date", "完工予定日", { type: "date" })}
              {field("work_period", "工事希望時期", { placeholder: "例: 2026年12月〜" })}
              {field("due_date", "見積提出期限", { type: "date" })}
            </div>
            <div className="split">
              <div>
                <h3>構造</h3>
                <div className="seg">
                  {(Object.keys(STRUCT) as Structure[]).map(k => (
                    <label key={k}><input type="radio" name="structure" checked={structure === k} onChange={() => setStructure(k)} />{STRUCT[k]}</label>
                  ))}
                </div>
              </div>
              <div>
                <h3>階数</h3>
                <div className="seg">
                  {([1, 2, 3] as Floors[]).map(k => (
                    <label key={k}><input type="radio" name="floors" checked={floors === k} onChange={() => setFloors(k)} />{FLOOR_LABEL[k]}</label>
                  ))}
                </div>
              </div>
            </div>
            <p className="hint">構造と階数に合わせて、工事内訳が絞り込まれます。数量は、内訳の表で直接入力します。</p>
            <label className="wide block">備考・特記事項（お取引先への伝達事項）
              <textarea rows={3} value={prop.remarks} onChange={setP("remarks")} placeholder="搬入経路、工事時間の制限、既存設備の状況など" />
            </label>
          </section>

          <section className="card" id="works">
            <div className="card-h"><h2>工事の選択</h2></div>
            <div className="fmt">
              <div className="plumb-h"><b>工事の種類</b><span>選択した種類に応じて、Excelの出力形式が替わります。</span></div>
              <div className="seg">
                {(Object.keys(FORMATS) as ExcelFormat[]).map(k => (
                  <button key={k} type="button" className={k === format ? "on" : ""} onClick={() => changeFormat(k)}>{FORMATS[k].name}</button>
                ))}
              </div>
            </div>

            {showPlumb && (
              <div className="plumb">
                <div className="plumb-h"><b>水回りの範囲</b><span>選んだ設備の工事だけが、給排水設備工事・住宅設備工事などの内訳に出ます</span></div>
                <h4 className="sub">設備を個別に選ぶ</h4>
                <div className="fx">
                  {plumbing.fixtures.map(f => (
                    <label key={f.key}><input type="checkbox" checked={fixturesOn.has(f.key)}
                      onChange={e => setFixturesOn(s => toggle(s, f.key, e.target.checked))} />{f.name}</label>
                  ))}
                </div>
              </div>
            )}

            <div className="cats">
              {loadError && <p className="empty">工事項目を読み込めませんでした。バックエンドの起動と NEXT_PUBLIC_API_BASE を確認してください。</p>}
              {catalog.map(c => (
                <label key={c.id} className={"cat " + (selected.has(c.id) ? "on" : "")}>
                  <input type="checkbox" checked={selected.has(c.id)} onChange={e => setSelected(s => toggle(s, c.id, e.target.checked))} />
                  <span className="ico"><svg viewBox="0 0 24 24" aria-hidden="true"><path d={ICONS[c.id] || ICONS[12]} /></svg></span>
                  {c.name}<small>{visibleItems(c, cond).length}項目</small>
                </label>
              ))}
            </div>
          </section>

          <section className="card" id="detail">
            <div className="card-h"><h2>内訳の確認</h2></div>
            {cats.length === 0 ? <p className="empty">「工事の選択」で工事を選ぶと、内訳がここに表示されます。</p> : cats.map(c => (
              <details key={c.id} open={openCats.has(c.id)} onToggle={e => {
                const open = (e.currentTarget as HTMLDetailsElement).open;
                setOpenCats(s => (s.has(c.id) === open ? s : toggle(s, c.id, open)));
              }}>
                <summary><span className="no">{c.id}</span>{c.name}<small>{visibleItems(c, cond).length}項目</small></summary>
                <div className="tbl"><table>
                  <thead><tr><th>名称</th><th>仕様</th><th>数量</th><th>単位</th><th>摘要</th><th></th></tr></thead>
                  <tbody>
                    {c.items.map((it, i) => !visible(it, cond) ? null : (
                      <tr key={i} className={[it.group ? "group" : "", needsQty(it) ? "todo" : ""].join(" ")}>
                        <td><input value={it.name} onChange={e => patchItem(c.id, i, { name: e.target.value })} /></td>
                        {it.group ? <td colSpan={4}></td> : <>
                          <td><input placeholder="仕様を記入" value={it.spec ?? ""} onChange={e => patchItem(c.id, i, { spec: e.target.value })} /></td>
                          <td className="q"><input type="number" step="any" inputMode="decimal" placeholder="数量" value={it.qty ?? ""}
                            onChange={e => patchItem(c.id, i, { qty: e.target.value === "" ? null : +e.target.value })} /></td>
                          <td className="u"><input value={it.unit ?? ""} onChange={e => patchItem(c.id, i, { unit: e.target.value })} /></td>
                          <td><input placeholder="摘要" value={it.note ?? ""} onChange={e => patchItem(c.id, i, { note: e.target.value })} /></td>
                        </>}
                        <td className="ac">
                          <button className="mini" title="削除" onClick={() => patchCat(c.id, its => its.filter((_, k) => k !== i))}>✕</button>
                          <button className="mini" title="上へ" onClick={() => moveUp(c.id, i)}>↑</button>
                          <button className="mini" title="下へ" onClick={() => moveDown(c.id, i)}>↓</button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table></div>
                <div className="tools">
                  <button className="mini" onClick={() => patchCat(c.id, its => [...its, { name: "", spec: "", qty: null, unit: "式" }])}>＋ 項目を追加</button>
                  <button className="mini" onClick={() => patchCat(c.id, its => [...its, { name: "【】", group: true }])}>＋ 見出し(【 】)を追加</button>
                </div>
              </details>
            ))}
          </section>
        </main>
      </div>

      <aside className="summary">
        <div className="sum-body">
          <h4>出力の内容</h4>
          <dl>
            <div><dt>構造・階数</dt><dd>{structLabel}</dd></div>
            <div><dt>工事の種類</dt><dd>{FORMATS[format].name}</dd></div>
            {format === "mizumawari" && <div><dt>水回りの設備</dt><dd>{plLabel || "なし"}</dd></div>}
            <div><dt>工事</dt><dd>{cats.length}工事</dd></div>
            <div><dt>内訳項目</dt><dd>{items.length}項目</dd></div>
            <div><dt>数量が未入力</dt><dd className={todo ? "warn" : "ok"}>{todo ? `${todo}項目` : "なし"}</dd></div>
          </dl>
        </div>
        <div className="sum-actions">
          <span id="status">{msg || `${cats.length}工事 / ${items.length}項目 を出力`}</span>
          <button className="ghost" title="今の項目構成を次回以降のひな形として保存" onClick={saveCatalog}>ひな形として保存</button>
          <button className="primary" disabled={busy} onClick={doExport}>{busy ? "作成中…" : "Excelを出力"}</button>
        </div>
      </aside>
    </div>
  );
}
