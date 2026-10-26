# 見積依頼書ジェネレーター

物件情報（構造・階数・面積）と工事タイプから、**単価欄が空のExcel見積依頼書**を作るツールです。
元フォーマットは「リフォーム・リノベーション工事見積書」。業者は黄色の単価欄に入力するだけで、金額・小計・合計が計算されます。

同じ内容を3つの形で用意しています。

| 形 | 場所 | 使いどころ |
|---|---|---|
| Webアプリ（FastAPI + HTML） | `backend/` `frontend/` | 自社サーバーで運用する本番用 |
| 単体HTML（サーバー不要） | `web/dist/index.html` | ネットに公開して見せる・社外に共有する |
| Excel版（マクロなし） | `excel/見積依頼書ジェネレーター.xlsx` | Excelだけで完結させたいとき |

DB（RDB）は使いません。標準の工事項目は `backend/catalog.json` のファイル1つで管理します。

## ディレクトリ構成

```
estimate_app/
├─ backend/              FastAPI（API とExcel生成）
│  ├─ main.py            API: GET/PUT /api/catalog, POST /api/export
│  ├─ excel_export.py    見積依頼書(.xlsx)の生成
│  ├─ catalog.json       工事・項目・工事タイプのひな形（全版で共通）
│  └─ requirements.txt
├─ frontend/             画面（静的ファイルのみ。ビルド不要）
│  ├─ index.html  style.css  app.js
│  └─ config.js          バックエンドのURL（API_BASE）を指定
├─ web/                  単体HTML版（サーバー不要）
│  ├─ build.py           frontend/ と catalog.json から web/dist/ を生成
│  ├─ export_xlsx.js     ブラウザ内のExcel生成（ExcelJS）
│  └─ dist/index.html    ← 静的ホスティングにそのまま置ける
├─ excel/                Excel版
│  ├─ build_excel.py     catalog.json から .xlsx を生成
│  └─ 見積依頼書ジェネレーター.xlsx
└─ start.bat             手元で起動（Windows）
```

`catalog.json` を変えたら、`python web/build.py` と `python excel/build_excel.py` を実行すると、単体HTML版とExcel版にも反映されます。

## 1. Webアプリ（front と back を分けて使う）

初回だけ依存を入れます。

```bash
pip install -r backend/requirements.txt
```

### 手元で一緒に動かす（いちばん簡単）

`start.bat` をダブルクリック。バックエンドが `frontend/` も配信し、ブラウザが開きます。

### front と back を別々に動かす

バックエンド（API だけ）:

```bash
cd backend
SERVE_FRONTEND=0 ALLOWED_ORIGINS=http://127.0.0.1:5500 python -m uvicorn main:app --port 8000
```

フロント（任意の静的ホスティング。ここでは確認用に簡易サーバー）:

```bash
cd frontend
python -m http.server 5500
```

`frontend/config.js` の `API_BASE` を `"http://127.0.0.1:8000"` にします。本番では、フロントを置くURLを `ALLOWED_ORIGINS` に、バックエンドのURLを `API_BASE` に指定してください。

### バックエンドの設定（環境変数）

| 名前 | 既定値 | 内容 |
|---|---|---|
| `SERVE_FRONTEND` | `1` | `0` にすると API だけを配信する |
| `ALLOWED_ORIGINS` | `http://127.0.0.1:5500,http://localhost:5500` | CORSで許可するフロントのURL（カンマ区切り） |
| `ENABLE_CATALOG_WRITE` | `0` | `1` にすると「ひな形として保存」を有効にする（`start.bat` では `1`） |
| `ENABLE_DOCS` | `0` | `1` にすると API ドキュメント（`/docs`）を表示する |

公開サーバーに置く場合の注意: 認証は付けていません。社外に公開するなら、`ENABLE_CATALOG_WRITE` は `0` のままにし、リバースプロキシやBasic認証で保護してください。ひな形の保存先は `catalog.json` のファイルなので、コンテナ等では再起動で消える環境があります。

## 2. 単体HTML版（ネットで見せる）

```bash
python web/build.py
```

`web/dist/index.html` 1ファイルができます。GitHub には `main` ブランチへ push すると `.github/workflows/pages.yml` が自動で GitHub Pages に公開します（初回はリポジトリの Settings > Pages で Source を「GitHub Actions」にします）。ほかにも Netlify、社内の共有サーバーなど、静的ファイルを置ける場所にそのまま置けば動きます。Excelの生成はブラウザ内で行うため、バックエンドは要りません（ExcelJS を CDN から読み込みます）。スマホ幅にも対応しています。

保存機能（ひな形の保存）はこの版にはありません。

## 3. Excel版

```bash
python excel/build_excel.py
```

マクロは使っていません。「入力」シートに STEP 1〜5 の順で入力・選択すると、「見積書」シートの内訳が数式で組み上がります。詳しくはブック内の「使い方」シートを見てください。

注意:
- 「見積書」シートは数式で行が決まります。業者に送るファイルを保存したあとは、自社側の「入力」シートを変えないでください。数式のない見積書にしたいときは、値と元の書式で別ブックに貼り付けます。
- 数式の計算ロジックは Python の計算エンジンで確認していますが、Excel 本体では未確認です。最初に一度、Excelで開いて表示を確かめてください。

## 数量の自動入力のしくみ

`catalog.json` の項目に `basis`（元になる数値）と `factor`（係数）を持たせています。例: 「材木費」は `basis: tsubo`, `factor: 1` で、延床坪数がそのまま数量になります。`minFloors`（最低階数）と `structures`（対象構造）で、階数・構造による絞り込みを行います。

## 補足

- 元のExcelにない項目（「軽量鉄骨下地」「コンクリート斫り・コア抜き」）は、鉄骨・RC向けに仮で追加したものです。名称・単位は実務に合わせて直してください。
- 標準項目は、元のExcelの「原価見積もり入力」シートにある案件の内訳です。特定の商品名が入っています。
