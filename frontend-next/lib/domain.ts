import type { Category, ExcelFormat, Item, Plumbing, Structure, Floors } from "./types";

export const API = (process.env.NEXT_PUBLIC_API_BASE ?? "").replace(/\/$/, "");
export const STRUCT: Record<Structure, string> = { wood: "木造", steel: "鉄骨造", rc: "RC造" };
export const FLOOR_LABEL: Record<Floors, string> = { 1: "平屋", 2: "2階建て", 3: "3階建て" };
// Excelの形式(見本のExcelに合わせた2種類)。工事名の初期値も形式ごとに持つ
export const FORMATS: Record<ExcelFormat, { name: string; title: string }> = {
  renovation: { name: "リノベーション工事", title: "リノベーション工事" },
  mizumawari: { name: "水回りリフォーム", title: "水回りリフォーム工事" },
};

export type Cond = { structure: Structure; floors: Floors; fixturesOn: Set<string> };

// 構造・階数・水回りの範囲に合う項目だけを表示・出力する
export const visible = (it: Item, c: Cond) =>
  (!it.minFloors || it.minFloors <= c.floors) &&
  (!it.structures || it.structures.includes(c.structure)) &&
  (!it.fixtures || it.fixtures.some(f => c.fixturesOn.has(f)));
export const visibleItems = (cat: Category, c: Cond) => cat.items.filter(i => !i.group && visible(i, c));
export const needsQty = (it: Item) => !it.group && (it.qty == null || (it.qty as unknown) === "");
export const sameSet = (a: Set<unknown>, b: Set<unknown>) => a.size === b.size && [...a].every(x => b.has(x));

export const hasFixtureItems = (cats: Category[], floors: Floors) =>
  cats.some(c => c.items.some(i => i.fixtures && (!i.minFloors || i.minFloors <= floors)));
export const plumbingLabel = (cats: Category[], pl: Plumbing, c: Cond) => {
  if (!hasFixtureItems(cats, c.floors)) return "";
  const names = pl.fixtures.filter(f => c.fixturesOn.has(f.key)).map(f => f.name);
  return names.length ? names.join("＋") : "なし";
};

export const ICONS: Record<number, string> = {
  1: "M14 4l6 6-3 3-6-6zM12 8L4 16l4 4 8-8",
  2: "M3 8l9-4 9 4-9 4zM3 13l9 4 9-4M3 17l9 4 9-4",
  3: "M4 20V8l8-4 8 4v12M9 20v-6h6v6M4 12h16",
  4: "M2 12l10-8 10 8M5 10v10h14V10",
  5: "M4 4h16v16H4zM4 12h16M12 4v8M8 12v8M16 12v8",
  6: "M5 3h14v18H5zM12 3v18M5 12h14",
  7: "M4 4h12v6H4zM16 7h4v6h-8v3M10 16h4v5h-4z",
  8: "M13 2L5 14h6l-1 8 8-12h-6z",
  9: "M12 3s6 6.5 6 11a6 6 0 0 1-12 0c0-4.5 6-11 6-11z",
  10: "M3 12h18v3a5 5 0 0 1-5 5H8a5 5 0 0 1-5-5zM6 12V6a2 2 0 0 1 4 0M5 20l-1 2M19 20l1 2",
  11: "M5 3v18M19 3v18M5 7h14M5 12h14M5 17h14",
  12: "M6 3h12v18l-3-2-3 2-3-2-3 2zM9 8h6M9 12h6",
};
