import json
import os
from pathlib import Path
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

# 公開環境では API ドキュメント(/docs 等)を出さない。手元で見たいときは ENABLE_DOCS=1。
_docs = os.getenv("ENABLE_DOCS", "0") == "1"
app = FastAPI(title="見積依頼書ジェネレーター", docs_url="/docs" if _docs else None,
              redoc_url=None, openapi_url="/openapi.json" if _docs else None)

# フロントを別ホスティングする場合に許可するオリジン(カンマ区切り)。例: ALLOWED_ORIGINS=https://example.github.io
origins = [o.strip() for o in os.getenv("ALLOWED_ORIGINS", "http://127.0.0.1:5500,http://localhost:5500").split(",") if o.strip()]
app.add_middleware(CORSMiddleware, allow_origins=origins, allow_methods=["GET", "POST", "PUT"],
                   allow_headers=["Content-Type"], expose_headers=["Content-Disposition"])


@app.middleware("http")
async def security_headers(request: Request, call_next):
    res = await call_next(request)
    res.headers.setdefault("X-Content-Type-Options", "nosniff")
    res.headers.setdefault("X-Frame-Options", "DENY")
    res.headers.setdefault("Referrer-Policy", "strict-origin-when-cross-origin")
    return res


# 入力の上限(巨大なリクエストでメモリやCPUを使い切られないようにする)
Text = Field(default="", max_length=200)
LongText = Field(default="", max_length=2000)
Num = Field(default=None, ge=0, le=1_000_000)
MAX_CATEGORIES = 50
MAX_ITEMS = 500


class Item(BaseModel):
    name: str = Field(max_length=200)
    spec: str = Text
    qty: float | None = Num
    unit: str = Field(default="", max_length=20)
    note: str = Text
    group: bool = False


class Category(BaseModel):
    name: str = Field(max_length=100)
    items: list[Item] = Field(max_length=MAX_ITEMS)


class Property(BaseModel):
    customer_name: str = Text
    work_title: str = Field(default="リノベーション工事", max_length=200)
    address: str = Text
    structure: str = Text
    company_name: str = Text
    vendor_name: str = Text
    due_date: str = Field(default="", max_length=20)
    work_period: str = Text
    total_area: float | None = Num
    tsubo: float | None = Num
    floor1_area: float | None = Num
    floor2_area: float | None = Num
    floor3_area: float | None = Num
    floors: int = Field(default=2, ge=1, le=3)
    roof_area: float | None = Num
    wall_area: float | None = Num
    scaffold_area: float | None = Num
    toilet_count: float | None = Num
    remarks: str = LongText


class ExportRequest(BaseModel):
    property: Property
    categories: list[Category] = Field(min_length=1, max_length=MAX_CATEGORIES)


# ---- ひな形(catalog.json)の保存用。想定外のキーや型は受け付けない ----
class CatalogItem(BaseModel):
    model_config = ConfigDict(extra="forbid")
    name: str = Field(max_length=200)
    spec: str = Text
    qty: float | None = Num
    unit: str = Field(default="", max_length=20)
    note: str = Text
    group: bool = False
    basis: str | None = Field(default=None, pattern=r"^[a-z0-9_]{1,40}$")
    factor: float | None = Field(default=None, ge=0, le=1000)
    minFloors: int | None = Field(default=None, ge=1, le=3)
    structures: list[str] | None = Field(default=None, max_length=3)


class CatalogCategory(BaseModel):
    model_config = ConfigDict(extra="forbid")
    id: int = Field(ge=1, le=10_000)
    name: str = Field(max_length=100)
    items: list[CatalogItem] = Field(max_length=MAX_ITEMS)
    summary_only: bool | None = None


class Preset(BaseModel):
    model_config = ConfigDict(extra="forbid")
    name: str = Field(max_length=100)
    ids: list[int] = Field(max_length=MAX_CATEGORIES)


class Catalog(BaseModel):
    model_config = ConfigDict(extra="forbid")
    categories: list[CatalogCategory] = Field(min_length=1, max_length=MAX_CATEGORIES)
    presets: list[Preset] = Field(default=[], max_length=50)


@app.get("/api/catalog")
def get_catalog():
    return json.loads(CATALOG_PATH.read_text(encoding="utf-8"))


@app.put("/api/catalog")
def save_catalog(body: Catalog):
    """画面で編集した標準項目を、次回以降のひな形として保存する。

    認証がないため既定では無効。手元で使うときだけ ENABLE_CATALOG_WRITE=1 にする(start.bat で指定済み)。
    """
    if os.getenv("ENABLE_CATALOG_WRITE", "0") != "1":
        raise HTTPException(status_code=403, detail="ひな形の保存は無効です")
    data = body.model_dump(exclude_unset=True)
    # 書き込み途中で落ちても catalog.json が壊れないよう、一時ファイル経由で置き換える
    tmp = CATALOG_PATH.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    tmp.replace(CATALOG_PATH)
    return {"ok": True}


@app.post("/api/export")
def export_excel(req: ExportRequest):
    buf = build_workbook(req.model_dump())
    p = req.property
    parts = [p.customer_name or "お客様", p.work_title, "見積依頼書"]
    if p.vendor_name:
        parts.append(p.vendor_name)
    filename = "_".join(parts) + ".xlsx"
    return Response(
        buf.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(filename)}"},
    )


# 手元で動かすときの利便性のため、既定ではフロントも同じサーバーから配信する。
# バックエンドだけを動かす場合は SERVE_FRONTEND=0 を指定する。
if os.getenv("SERVE_FRONTEND", "1") == "1" and FRONTEND.exists():
    app.mount("/", StaticFiles(directory=FRONTEND, html=True), name="frontend")
