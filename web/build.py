"""ネット公開用のサンプル(単一HTML)を生成する。

FastAPIの代わりに、カタログを埋め込み、Excel出力をブラウザ内(ExcelJS)で行う。
実行: python web/build.py  ->  web/dist/index.html (静的ホスティング用) と artifact.html (claude.ai用)
"""
import json
import re
from pathlib import Path

HERE = Path(__file__).parent
ROOT = HERE.parent  # estimate_app/
catalog = (ROOT / "backend" / "catalog.json").read_text(encoding="utf-8")
html = (ROOT / "frontend" / "index.html").read_text(encoding="utf-8")
css = (ROOT / "frontend" / "style.css").read_text(encoding="utf-8")
js = (ROOT / "frontend" / "app.js").read_text(encoding="utf-8")

# 1) カタログは埋め込み
# (サーバーから読み込む部分を、埋め込んだ CATALOG に差し替える)
start = js.index("  let cat;\n  try {")
end = js.index("    return;\n  }\n", start) + len("    return;\n  }\n")
js = js[:start] + "  const cat = CATALOG;\n" + js[end:]

# 2) 出力・ひな形保存のハンドラをサンプル用に置換
start = js.index('$("#export").addEventListener')
end = js.index("// ---- 起動")
js = js[:start] + (HERE / "export_xlsx.js").read_text(encoding="utf-8") + "\n" + js[end:]

body = re.search(r"<body>(.*)</body>", html, re.S).group(1)
body = re.sub(r'<script src="(app|config).js"></script>\s*', "", body)
body = re.sub(r'<button id="saveCatalog".*?</button>', "", body)
body = body.replace('<p class="banner" id="sampleNote" hidden>', '<p class="banner" id="sampleNote">')

extra_css = """
.app{min-height:100vh}
@media (max-width:760px){.top,main,.banner,.side{padding-left:max(16px,env(safe-area-inset-left,0px));padding-right:max(16px,env(safe-area-inset-right,0px))}.banner{margin-left:0;margin-right:0;border-radius:0}}
@media (max-width:1180px){.summary{padding-left:max(16px,env(safe-area-inset-left,0px));padding-right:max(16px,env(safe-area-inset-right,0px))}}
"""

seed = """
// 例として入れておく数値(サンプル)
for (const [k, v] of Object.entries({ floor1_area: 63.6, floor2_area: 33, roof_area: 83, wall_area: 123.3, scaffold_area: 360, toilet_count: 2, cloth_len: 670 })) {
  const el = document.querySelector(`[data-n=${k}]`);
  if (el) { el.value = v; el.dispatchEvent(new Event("input", { bubbles: true })); }
}
"""

out = f"""<title>見積依頼書ジェネレーター</title>
<meta name="claude:viewport-fit-sides" content="cover">
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;500;700&display=swap">
<style>
{css}
{extra_css}
</style>
<script src="https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js" integrity="sha512-dlPw+ytv/6JyepmelABrgeYgHI0O+frEwgfnPdXDTOIZz+eDgfW07QXG02/O8COfivBdGNINy+Vex+lYmJ5rxw==" crossorigin="anonymous" referrerpolicy="no-referrer"></script>
{body}
<script>
const CATALOG = {catalog};
{js}
{seed}
</script>
"""
(HERE / "dist").mkdir(exist_ok=True)
# claude.ai の Artifact 用(ページ側で doctype/head が付与されるため、断片のまま)
(HERE / "dist" / "artifact.html").write_text(out, encoding="utf-8")
# 通常のWebサイト用(静的ホスティングにそのまま置ける完全なHTML)
meta = '<meta name="claude:viewport-fit-sides" content="cover">\n'
page = out.replace(meta, "")
head_end = page.index("</style>") + len("</style>")
standalone = (
    '<!doctype html>\n<html lang="ja">\n<head>\n<meta charset="utf-8">\n'
    '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n'
    + page[:head_end] + "\n</head>\n<body>" + page[head_end:] + "</body>\n</html>\n"
)
(HERE / "dist" / "index.html").write_text(standalone, encoding="utf-8")
print("ok", len(out))
