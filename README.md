# TSUMIKI

物件情報（構造・階数）と、工事・水回りの範囲から、**単価欄が空のExcel見積依頼書**を作るツールです。数量は案件ごとに内訳の表で手入力します。
元フォーマットは「リフォーム・リノベーション工事見積書」。業者は黄色の単価欄に入力するだけで、金額・小計・合計が計算されます。

同じ内容を4つの形で用意しています。

| 形 | 場所 | 使いどころ |
|---|---|---|
| Webアプリ（FastAPI + HTML） | `backend/` `frontend/` | 自社サーバーで運用する本番用 |
| Next.js フロント（npm管理） | `frontend-next/` | `frontend/` と同じ画面を Next.js(React) で作り直したもの。FastAPI と組み合わせて使う |
| 単体HTML（サーバー不要） | `web/dist/index.html` | ネットに公開して見せる・社外に共有する |
| Excel版（マクロなし） | `excel/TSUMIKI.xlsx` | Excelだけで完結させたいとき（旧仕様のまま。下の注意を参照） |

DB（RDB）は使いません。標準の工事項目は `backend/catalog.json` のファイル1つで管理します。

## ディレクトリ構成

```
estimate_app/
├─ backend/              FastAPI（API とExcel生成）
│  ├─ main.py            API: GET/PUT /api/catalog, POST /api/export
│  ├─ excel_export.py    見積依頼書(.xlsx)の生成
│  ├─ catalog.json       工事・項目・工事の種類のひな形（全版で共通）
│  └─ requirements.txt
├─ frontend/             画面（静的ファイルのみ。ビルド不要）
│  ├─ index.html  style.css  app.js
│  └─ config.js          バックエンドのURL（API_BASE）を指定
├─ frontend-next/        Next.js 版フロント（package.json / package-lock.json で版を管理）
├─ web/                  単体HTML版（サーバー不要）
│  ├─ build.py           frontend/ と catalog.json から web/dist/ を生成
│  ├─ export_xlsx.js     ブラウザ内のExcel生成（ExcelJS）
│  └─ dist/index.html    ← 静的ホスティングにそのまま置ける
├─ excel/                Excel版
│  ├─ build_excel.py     catalog.json から .xlsx を生成
│  └─ TSUMIKI.xlsx
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

## 1-2. Next.js フロント（`frontend-next/`）

Node.js 20.9 以上が必要です。依存の版は `package.json` で固定し（`.npmrc` の `save-exact`）、`package-lock.json` で再現できます。

```bash
cd frontend-next
npm ci
npm run dev
```

`http://localhost:3000` を開きます。バックエンドは別に起動し（上の「front と back を別々に動かす」）、`ALLOWED_ORIGINS` に `http://localhost:3000` を入れます。バックエンドのURLは `.env.local` の `NEXT_PUBLIC_API_BASE`（見本は `.env.example`）で指定します。

```bash
npm run build      # out/ に静的ファイルを書き出す（Netlify等にそのまま置ける）
npm run typecheck  # 型チェック
```

版の管理は npm の semver で行います。リリース時に `npm run release:patch`（または `release:minor` / `release:major`）を実行すると、`package.json` の `version` が上がります。Git リポジトリ内では、コミットとタグも自動で作られます。

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

注意（Excel版）: Excel版は、面積から数量を自動で入れる旧仕様のままです。画面（Webアプリ・単体HTML）の「数量は手入力」「水回りの範囲」は、Excel版にはまだ入っていません。

## 工事の種類とExcelの出力形式

画面の「工事の選択」で、工事の種類（リノベーション工事・水回りリフォーム・外装工事・内装工事）を選びます。選ぶと、その種類の大項目が選ばれ、Excelの形も切り替わります（`catalog.json` の `presets` の `name` / `ids` / `format`）。Excelは、見本のExcelに合わせた2つの形です。

| 形式 | 見本 | シート |
|---|---|---|
| `renovation` リノベーション工事、`exterior` 外装工事、`interior` 内装工事 | リノベーション工事_見積依頼書.xlsx | 見積書（8列）、物件情報 |
| `mizumawari` 水回りリフォーム | 水回りリフォーム工事 見積書.xlsx | 表紙、見積もり内訳（11列）、物件情報 |

- どちらも、全セルが細い実線の罫線です。見本にならい、単価・摘要の欄に色は付けていません。
- 水回りリフォームの「表紙」は、御見積金額が「見積もり内訳」の合計を参照します（単価が入ると金額が出ます）。見本にある原価内訳のシートは、社内の原価と掛率が入る内部用のため、業者へ渡すExcelには含めていません。
- 作成は `backend/excel_export.py`（Webアプリ）と `web/export_xlsx.js`（単体HTML）で、レイアウトは同じです。変えるときは両方をそろえてください。
## 画面の仕様（Webアプリ・単体HTML）

- **工事の種類**: リノベーション工事（すべての大項目）／水回りリフォーム（解体・内外装・電気設備・給排水設備・住宅設備・諸経費だけ）／外装工事／内装工事。種類ごとの大項目は `catalog.json` の `presets` で変えられます。
- **依頼元（自社名）**: 初期値は「株式会社石丸ハウスセンター」です。画面で書き換えられます。「提出先（業者名）」の入力欄はありません。
- **水回りの設備**: 「水回りリフォーム」を選んだときだけ、「設備を個別に選ぶ」（キッチン／風呂＆洗面器／トイレ）が出ます。選んだ設備に関係する項目だけが、給排水設備工事・住宅設備工事・解体工事・電気設備工事の内訳に出ます。ほかの種類では、設備はすべて対象です。
- **数量は手入力**: 面積などから自動で入れる機能はありません。数量が空の項目は内訳で黄色になり、右の「出力の内容」に件数が出ます。単位が「式」の項目だけ、初期値を 1 にしています。
- **構造・階数**: 構造（木造／鉄骨造／RC造）と階数（平屋／2階建て／3階建て）で、項目を絞り込みます。

## 項目の絞り込みのしくみ

`catalog.json` の項目に、表示の条件を持たせています。

| キー | 内容 |
|---|---|
| `minFloors` | この階数以上のときだけ表示（例: 2階の項目） |
| `structures` | 対象の構造（`wood` / `steel` / `rc`） |
| `fixtures` | 水回りの設備（`kitchen` / `bathwash` / `toilet`）。どれかが選ばれていれば表示 |

水回りの設備は、`catalog.json` の `plumbing` で定義しています。`basis` / `factor`（面積からの数量計算）は、Excel版だけが使います。

## 補足

- 元のExcelにない項目（「軽量鉄骨下地」「コンクリート斫り・コア抜き」）は、鉄骨・RC向けに仮で追加したものです。名称・単位は実務に合わせて直してください。
- 標準項目は、元のExcelの「原価見積もり入力」シートにある案件の内訳です。特定の商品名が入っています。

## セキュリティの確認（2026-11-14）

- **依存パッケージ**: `pip-audit`（`backend/requirements.txt`）と `npm audit`（`frontend-next`）で、既知の脆弱性は0件でした。
- **入力**: 文字数・件数・数値の範囲をバックエンドで検査しています。ひな形の保存は、想定外のキーを受け付けず、既定では無効です。
- **Excelへの数式の混入**: 名称や備考が「=」で始まっても、数式ではなく文字として書き込みます（バックエンド・ブラウザ版とも）。
- **画面（XSS）**: 画面に出す文字は、すべてエスケープしています（Next.js版は React が処理）。
- **ヘッダー**: `Content-Security-Policy`（自分のサーバーのスクリプトだけを実行）、`X-Frame-Options`、`X-Content-Type-Options` を付けています。`/api/` の応答は保存しない設定（`no-store`）です。
- **ファイル名**: 出力ファイル名の記号や制御文字は、置き換えています。
- **未対応**: 認証、利用回数の制限（レート制限）は入っていません。社外に公開するサーバーでは、リバースプロキシやBasic認証で保護してください。
- このプロジェクトでは PHP は使っていません（Python / JavaScript のみ）。

