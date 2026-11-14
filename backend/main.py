import json
import os
from pathlib import Path
from typing import Annotated, Literal
from urllib.parse import quote

from fastapi import FastAPI, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import Response
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, ConfigDict, Field

from excel_export import build_workbook

BASE = Path(__file__).parent
CATALOG_PATH = BASE / "catalog.json"
FRONTEND = BASE.parent / "frontend"


# ---- 設定(環境変数)。詳しくは README の「バックエンドの設定」 ----
def _flag(name: str, default: str) -> bool:
    return os.getenv(name, default) == "1"


SERVE_FRONTEND = _flag("SERVE_FRONTEND", "1")  # frontend/ も同じサーバーから配信する
ENABLE_DOCS = _flag("ENABLE_DOCS", "0")        # API ドキュメント(/docs)を出す。公開環境では出さない
# フロントを別ホスティングする場合に許可するオリジン(カンマ区切り)。例: https://example.github.io
ALLOWED_ORIGINS = [o.strip() for o in os.getenv("ALLOWED_ORIGINS", "http://127.0.0.1:5500,http://localhost:5500").split(",")
                   if o.strip()]


def catalog_write_enabled() -> bool:
    """ひな形の保存を許可するか。認証がないため既定は無効(start.bat では有効にしている)。"""
    return _flag("ENABLE_CATALOG_WRITE", "0")


app = FastAPI(title="TSUMIKI", docs_url="/docs" if ENABLE_DOCS else None,
              redoc_url=None, openapi_url="/openapi.json" if ENABLE_DOCS else None)
app.add_middleware(CORSMiddleware, allow_origins=ALLOWED_ORIGINS, allow_methods=["GET", "POST", "PUT"],
                   allow_headers=["Content-Type"], expose_headers=["Content-Disposition"])


# 画面は自分のサーバーのスクリプトだけを実行する(Google Fonts のみ外部から読む)。万一HTMLに不正な文字が混ざっても、実行させない
CSP = ("default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline' https://fonts.googleapis.com; "
       "font-src https://fonts.gstatic.com; img-src 'self' data:; connect-src 'self'; "
       "base-uri 'none'; form-action 'none'; frame-ancestors 'none'")


@app.middleware("http")
async def security_headers(request: Request, call_next):
    res = await call_next(request)
    res.headers.setdefault("X-Content-Type-Options", "nosniff")
    res.headers.setdefault("X-Frame-Options", "DENY")
    res.headers.setdefault("Referrer-Policy", "strict-origin-when-cross-origin")
    if request.url.path.startswith("/api/"):
        res.headers.setdefault("Cache-Control", "no-store")  # 見積の内容を、ブラウザや中継サーバーに残さない
    if not request.url.path.startswith(("/docs", "/openapi.json")):  # /docs は外部のスクリプトを使うので除く
        res.headers.setdefault("Content-Security-Policy", CSP)
    return res


# ---- 入力の形と上限 ----
# 上限は、巨大なリクエストでメモリやCPUを使い切られないようにするためのもの。
ShortText = Annotated[str, Field(max_length=200)]
LongText = Annotated[str, Field(max_length=2000)]
Unit = Annotated[str, Field(max_length=20)]
Amount = Annotated[float | None, Field(ge=0, le=1_000_000)]  # 面積・数量など
MAX_CATEGORIES = 50
MAX_ITEMS = 500
UNSAFE_FILENAME_CHARS = set('\\/:*?"<>|')


class Item(BaseModel):
    """内訳の1行。group=True は【 】の見出し行。"""
    name: ShortText
    spec: ShortText = ""
    qty: Amount = None
    unit: Unit = ""
    note: ShortText = ""
    group: bool = False


class Category(BaseModel):
    """大項目(解体工事、内装工事など)。"""
    name: Annotated[str, Field(max_length=100)]
    items: list[Item] = Field(max_length=MAX_ITEMS)


class Property(BaseModel):
    """物件情報。画面の「物件情報」欄に対応する。"""
    customer_name: ShortText = ""
    work_title: ShortText = "リノベーション工事"
    address: ShortText = ""
    structure: ShortText = ""
    company_name: ShortText = "株式会社石丸ハウスセンター"  # 自社名(初期値。画面で書き換え可)
    due_date: Annotated[str, Field(max_length=20)] = ""
    work_period: ShortText = ""
    floors: Annotated[int, Field(ge=1, le=3)] = 2
    plumbing: ShortText = ""  # 水回りの範囲(例: キッチン＋トイレ)
    remarks: LongText = ""
    format: Literal["renovation", "mizumawari", "exterior", "interior"] = "renovation"  # 工事の種類=Excelの形(リノベーション工事・水回りリフォーム・外装工事・内装工事)


class ExportRequest(BaseModel):
    property: Property
    categories: list[Category] = Field(min_length=1, max_length=MAX_CATEGORIES)


# ---- ひな形(catalog.json)。保存時は想定外のキーや型を受け付けない ----
class CatalogItem(Item):
    model_config = ConfigDict(extra="forbid")
    basis: Annotated[str, Field(pattern=r"^[a-z0-9_]{1,40}$")] | None = None  # 数量の元になる数値(例: tsubo)
    factor: Annotated[float, Field(ge=0, le=1000)] | None = None              # 数量 = basis × factor
    minFloors: Annotated[int, Field(ge=1, le=3)] | None = None                 # この階数以上で表示
    structures: Annotated[list[str], Field(max_length=3)] | None = None        # 対象の構造(wood/steel/rc)
    fixtures: Annotated[list[Annotated[str, Field(pattern=r"^[a-z0-9_]{1,40}$")]], Field(max_length=10)] | None = None  # 水回りの設備(kitchen 等)


class CatalogCategory(BaseModel):
    model_config = ConfigDict(extra="forbid")
    id: Annotated[int, Field(ge=1, le=10_000)]
    name: Annotated[str, Field(max_length=100)]
    items: list[CatalogItem] = Field(max_length=MAX_ITEMS)
    summary_only: bool | None = None


class Preset(BaseModel):
    """工事タイプ(工事・内装・外装)。ids は選択する大項目の id。"""
    model_config = ConfigDict(extra="forbid")
    name: Annotated[str, Field(max_length=100)]
    ids: list[int] = Field(max_length=MAX_CATEGORIES)
    format: Literal["renovation", "mizumawari", "exterior", "interior"] | None = None  # この種類(工事・内装・外装)のキー


class Fixture(BaseModel):
    """水回りの設備(キッチン、風呂＆洗面器、トイレ)。key は項目の fixtures で参照する。"""
    model_config = ConfigDict(extra="forbid")
    key: Annotated[str, Field(pattern=r"^[a-z0-9_]{1,40}$")]
    name: Annotated[str, Field(max_length=100)]


class PlumbingPattern(BaseModel):
    """水回りのよく使う組み合わせ(例: キッチン＋トイレ)。"""
    model_config = ConfigDict(extra="forbid")
    name: Annotated[str, Field(max_length=100)]
    fixtures: list[Annotated[str, Field(pattern=r"^[a-z0-9_]{1,40}$")]] = Field(max_length=10)


class Plumbing(BaseModel):
    model_config = ConfigDict(extra="forbid")
    fixtures: list[Fixture] = Field(max_length=10)
    patterns: list[PlumbingPattern] = Field(default_factory=list, max_length=20)  # 画面では使わない(旧版の名残)


class Catalog(BaseModel):
    model_config = ConfigDict(extra="forbid")
    categories: list[CatalogCategory] = Field(min_length=1, max_length=MAX_CATEGORIES)
    presets: list[Preset] = Field(default=[], max_length=50)
    plumbing: Plumbing | None = None


# ---------------------------------------------------------------- API

@app.get("/api/catalog")
def get_catalog():
    return json.loads(CATALOG_PATH.read_text(encoding="utf-8"))


@app.put("/api/catalog")
def save_catalog(body: Catalog):
    """画面で編集した標準項目を、次回以降のひな形として保存する。"""
    if not catalog_write_enabled():
        raise HTTPException(status_code=403, detail="ひな形の保存は無効です")
    data = body.model_dump(exclude_unset=True)
    # 書き込み途中で落ちても catalog.json が壊れないよう、一時ファイル経由で置き換える
    tmp = CATALOG_PATH.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    tmp.replace(CATALOG_PATH)
    return {"ok": True}


@app.post("/api/export")
def export_excel(req: ExportRequest):
    """見積依頼書(.xlsx)を作って返す。"""
    buf = build_workbook(req.model_dump())
    return Response(
        buf.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(_export_filename(req.property))}"},
    )


def _export_filename(p: Property) -> str:
    """例: 山田太郎_リノベーション工事_見積依頼書_〇〇工務店.xlsx"""
    parts = [p.customer_name or "お客様", p.work_title, "見積依頼書"]
    # ファイル名に使えない文字(制御文字を含む)は全角の「＿」に置き換える
    name = "".join("＿" if ch in UNSAFE_FILENAME_CHARS or ord(ch) < 32 else ch
                   for ch in "_".join(x for x in parts if x))
    return name[:120] + ".xlsx"


# ---------------------------------------------------------------- 画面の配信

# 手元で動かすときの利便性のため、既定ではフロントも同じサーバーから配信する。
# バックエンドだけを動かす場合は SERVE_FRONTEND=0 を指定する。
if SERVE_FRONTEND and FRONTEND.exists():
    app.mount("/", StaticFiles(directory=FRONTEND, html=True), name="frontend")
