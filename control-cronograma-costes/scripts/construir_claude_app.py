"""Prepara la app de claude.ai en claude_app/publicar/ (los archivos que se publican como artefacto).

Copia el motor y el informe compartidos con la app sin conexión, las librerías (SheetJS y pdf.js,
licencia Apache-2.0) e inserta los estilos del informe en index.html.
"""
import re
import shutil
from pathlib import Path

D = Path(__file__).resolve().parent
RAIZ = D.parent / "claude_app"
SAL = RAIZ / "publicar"
SAL.mkdir(parents=True, exist_ok=True)

cabecera = (D / "app" / "cabecera.html").read_text(encoding="utf-8")
css = re.search(r"<style>(.*?)</style>", cabecera, re.S).group(1)
html = (RAIZ / "src" / "index.html").read_text(encoding="utf-8").replace("/*__CSS_INFORME__*/", css, 1)
(SAL / "index.html").write_text(html, encoding="utf-8")
for origen, destino in ((D / "app" / "motor.js", "motor.js"), (D / "app" / "informe.js", "informe.js"),
                        (RAIZ / "src" / "app.js", "app.js"), (D / "vendor" / "xlsx.full.min.js", "xlsx.full.min.js"),
                        (D / "vendor" / "pdfjs" / "pdf.min.js", "pdf.min.js"),
                        (D / "vendor" / "pdfjs" / "pdf.worker.min.js", "pdf.worker.min.js")):
    shutil.copyfile(origen, SAL / destino)
print("OK ->", SAL, sorted(p.name for p in SAL.iterdir()))
