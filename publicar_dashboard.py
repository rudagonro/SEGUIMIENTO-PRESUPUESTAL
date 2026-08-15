"""
PUBLICADOR DE DASHBOARD - SEGUIMIENTO PRESUPUESTAL BRIAV32
==========================================================
Copia el HTML actualizado como index.html y hace git push
para actualizar GitHub Pages.

USO:
  python publicar_dashboard.py

Ejecutar DESPUES de actualizar_dashboard.py
"""

import subprocess, sys, shutil, os

HTML  = "SEGUIMIENTO PRESUPUESTAL.html"
INDEX = "index.html"

# Verificar que el HTML existe
if not os.path.exists(HTML):
    print(f"ERROR: No se encontro '{HTML}'. Ejecuta primero actualizar_dashboard.py")
    sys.exit(1)

print(f"Copiando {HTML} -> {INDEX} ...")
shutil.copy2(HTML, INDEX)

def run(cmd):
    result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8")
    if result.stdout.strip(): print(result.stdout.strip())
    if result.stderr.strip(): print(result.stderr.strip())
    return result.returncode

print("Publicando en GitHub Pages ...")
run(["git", "add", INDEX])
run(["git", "commit", "-m", "Actualizacion datos dashboard"])
code = run(["git", "push"])

print()
if code == 0:
    print("=" * 50)
    print("  PUBLICADO CORRECTAMENTE")
    print("=" * 50)
    print("  URL: https://rudagonro.github.io/SEGUIMIENTO-PRESUPUESTAL/")
    print("  (puede tardar 1-2 minutos en reflejarse)")
    print("=" * 50)
else:
    print("ERROR al hacer push. Verifica tu conexion a internet.")
