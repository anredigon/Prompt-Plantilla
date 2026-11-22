"""Ensambla la app de escritorio en un único HTML autocontenido: ../App_Control_Obra.html

Une la cabecera/estilos (app/cabecera.html), la app (app/escritorio.html), el motor de cálculo (app/motor.js),
el informe (app/informe.js), utilidades (app/comun.js) y SheetJS (vendor/xlsx.full.min.js, Apache-2.0).
"""
from pathlib import Path

D = Path(__file__).resolve().parent
html = (D / "app" / "cabecera.html").read_text(encoding="utf-8") + (D / "app" / "escritorio.html").read_text(encoding="utf-8")
html = html.replace("<title>Control de cronograma y costes</title>", "<title>Control de obra</title>")
for marca, ruta in (("/*__XLSX__*/", D / "vendor" / "xlsx.full.min.js"), ("/*__MOTOR__*/", D / "app" / "motor.js"),
                    ("/*__INFORME__*/", D / "app" / "informe.js"), ("/*__COMUN__*/", D / "app" / "comun.js")):
    codigo = ruta.read_text(encoding="utf-8")
    assert "</script" not in codigo.lower(), f"{ruta.name} contiene </script>"
    html = html.replace(marca, codigo, 1)
salida = D.parent / "App_Control_Obra.html"
salida.write_text(html, encoding="utf-8")
print(f"OK -> {salida} ({salida.stat().st_size / 1024:.0f} KB)")
