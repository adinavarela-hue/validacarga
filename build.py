"""Genera dist/ValidaCarga.html: la app completa en un solo archivo (se abre con doble clic, sin carpetas)."""
import pathlib

RAIZ = pathlib.Path(__file__).parent
SCRIPTS = ["validador.js", "perfiles.js", "plantilla.js", "demo-plantilla.js", "app.js"]

html = (RAIZ / "index.html").read_text(encoding="utf-8")
css = (RAIZ / "css" / "estilos.css").read_text(encoding="utf-8")
html = html.replace('<link rel="stylesheet" href="css/estilos.css">', f"<style>\n{css}\n</style>")
for nombre in SCRIPTS:
    js = (RAIZ / "js" / nombre).read_text(encoding="utf-8").replace("</script", "<\\/script")
    html = html.replace(f'<script src="js/{nombre}"></script>', f"<script>\n{js}\n</script>")
assert 'src="js/' not in html and "css/estilos" not in html, "Quedó una referencia a archivo externo"
salida = RAIZ / "dist" / "ValidaCarga.html"
salida.parent.mkdir(exist_ok=True)
salida.write_text(html, encoding="utf-8")
print(f"Listo: {salida} ({salida.stat().st_size // 1024} KB)")
