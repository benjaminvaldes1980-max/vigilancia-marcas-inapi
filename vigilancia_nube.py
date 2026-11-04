# -*- coding: utf-8 -*-
"""
Ejecuta la vigilancia de marcas INAPI en la nube (GitHub Actions).

Usa el mismo motor que la versión de escritorio (vigilancia_inapi.py) y toma
los datos del correo desde los "secrets" del repositorio:
    GMAIL_USUARIO     cuenta de Gmail que envía (ej.: tucuenta@gmail.com)
    GMAIL_CLAVE_APP   contraseña de aplicación de 16 letras
    AVISO_PARA        (opcional) a quién avisar; varios separados por coma
Las marcas se editan en palabras_clave.txt.
"""
import base64
import os
import shutil
import sys
from pathlib import Path

import vigilancia_inapi as v

RAIZ = Path(__file__).resolve().parent
DOCS = RAIZ / "docs"

# En la nube no hay ventanas: los avisos se escriben en el registro
v.aviso = lambda titulo, texto, preguntar=False: (print(f"[{titulo}] {texto}"), False)[1]
v.abrir = lambda ruta: None
v.mostrar_en_explorador = lambda ruta: None

_cargar_config = v.cargar_config


def config_nube():
    conf = _cargar_config()
    conf["una_palabra"] = os.environ.get("UNA_PALABRA", "si").strip().lower() not in ("no", "false", "0")
    usuario = os.environ.get("GMAIL_USUARIO", "").strip()
    clave = os.environ.get("GMAIL_CLAVE_APP", "").replace(" ", "").strip()
    para = os.environ.get("AVISO_PARA", "").strip() or usuario
    conf["correo"].update({
        "activo": bool(usuario and clave),
        "remitente": usuario,
        "destinatarios": para,
        "clave": ("b64:" + base64.b64encode(clave.encode()).decode()) if clave else "",
        "servidor": "smtp.gmail.com",
        "puerto": 465,
        "resumen_sin_novedades": os.environ.get("RESUMEN_DIARIO", "no").strip().lower() in ("si", "sí", "true", "1"),
    })
    return conf


v.cargar_config = config_nube

# Los PDF no se guardan en el repositorio: el informe enlaza solo a INAPI
_construir_informe = v.construir_informe


def informe_nube(resumen):
    for c in resumen.get("nuevas", []):
        c["archivo"] = ""
    return _construir_informe(resumen)


v.construir_informe = informe_nube


def publicar():
    """Deja el último informe y el historial en docs/ (página pública de GitHub Pages)."""
    DOCS.mkdir(exist_ok=True)
    if v.ULTIMO_INFORME.exists():
        shutil.copyfile(v.ULTIMO_INFORME, DOCS / "index.html")
    elif not (DOCS / "index.html").exists():
        (DOCS / "index.html").write_text(
            "<!DOCTYPE html><meta charset='utf-8'><title>Vigilancia de marcas INAPI</title>"
            "<body style='font-family:Segoe UI,Arial;padding:40px'><h1>Vigilancia de marcas INAPI</h1>"
            "<p>Todavía no hay revisiones con documentos nuevos.</p></body>", encoding="utf-8")
    if v.HISTORIAL.exists():
        shutil.copyfile(v.HISTORIAL, DOCS / "historial_coincidencias.csv")
    if v.HIST_REVISIONES.exists():
        shutil.copyfile(v.HIST_REVISIONES, DOCS / "historial_revisiones.csv")
    (DOCS / ".nojekyll").write_text("", encoding="utf-8")
    # el registro no crece sin límite
    if v.REGISTRO.exists():
        lineas = v.REGISTRO.read_text(encoding="utf-8", errors="replace").splitlines()
        if len(lineas) > 3000:
            v.REGISTRO.write_text("\n".join(lineas[-3000:]) + "\n", encoding="utf-8")


if __name__ == "__main__":
    conf = config_nube()
    print(f"Marcas vigiladas: {', '.join(v.leer_lista_claves())}")
    print("Correo de aviso: " + ("activado → " + conf["correo"]["destinatarios"] if conf["correo"]["activo"]
                                 else "desactivado (faltan los secrets GMAIL_USUARIO / GMAIL_CLAVE_APP)"))
    sys.argv = [sys.argv[0], "--manual"]
    codigo = v.main()
    publicar()
    sys.exit(codigo)
