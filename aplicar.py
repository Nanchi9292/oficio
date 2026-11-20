#!/usr/bin/env python3
"""Uso: python aplicar.py oficio.html   ->   genera oficio_pulido.html (el original no se modifica)."""
import sys, pathlib
if len(sys.argv) != 2:
    sys.exit(__doc__)
src = pathlib.Path(sys.argv[1])
html = src.read_text(encoding="utf-8")
css = '<link rel="stylesheet" href="pulido.css">\n'
js = '<script src="mejoras_excel.js"></script>\n'
if "pulido.css" not in html:
    html = html.replace("</head>", css + "</head>", 1)
if "mejoras_excel.js" not in html:
    html = html.replace("</body>", js + "</body>", 1)
out = src.with_name(src.stem + "_pulido.html")
out.write_text(html, encoding="utf-8")
print("Listo:", out, "\nCopie pulido.css y mejoras_excel.js a la misma carpeta.")
