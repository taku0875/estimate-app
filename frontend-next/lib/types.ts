export type Item = {
  name: string; spec?: string; qty?: number | null; unit?: string; note?: string;
  group?: boolean; minFloors?: number; structures?: string[]; fixtures?: string[];
};
export type Category = { id: number; name: string; items: Item[] };
export type ExcelFormat = "renovation" | "mizumawari";
export type Preset = { name: string; ids: number[]; format?: ExcelFormat };
export type Fixture = { key: string; name: string };
export type Pattern = { name: string; fixtures: string[] };
export type Plumbing = { fixtures: Fixture[]; patterns: Pattern[] };
export type CatalogData = { categories: Category[]; presets?: Preset[]; plumbing?: Plumbing };
export type Structure = "wood" | "steel" | "rc";
export type Floors = 1 | 2 | 3;
export type Prop = {
  customer_name: string; work_title: string; address: string; company_name: string;
  vendor_name: string; work_period: string; due_date: string; remarks: string;
};
