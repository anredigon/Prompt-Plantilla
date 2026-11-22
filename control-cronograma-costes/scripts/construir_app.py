"""Ensambla la app en un único HTML autocontenido: ../App_Control_Proyecto.html

Une la cabecera/estilos (app/cabecera.html), el cuerpo (app/app_cuerpo.html), el motor de cálculo
(app/motor.js) y la librería SheetJS (vendor/xlsx.full.min.js, Apache-2.0) para leer Excel sin conexión.
"""
from pathlib import Path

D = Path(__file__).resolve().parent
html = (D / "app" / "cabecera.html").read_text(encoding="utf-8") + (D / "app" / "app_cuerpo.html").read_text(encoding="utf-8")
libreria = (D / "vendor" / "xlsx.full.min.js").read_text(encoding="utf-8")
motor = (D / "app" / "motor.js").read_text(encoding="utf-8")
for nombre, codigo in (("libreria", libreria), ("motor", motor)):
    assert "</script" not in codigo.lower(), f"{nombre} contiene </script>"
html = html.replace("/*__XLSX__*/", libreria, 1).replace("/*__MOTOR__*/", motor, 1)
salida = D.parent / "App_Control_Proyecto.html"
salida.write_text(html, encoding="utf-8")
print(f"OK -> {salida} ({salida.stat().st_size / 1024:.0f} KB)")
