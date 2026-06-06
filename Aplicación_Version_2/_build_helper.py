"""
Helper de compilacion: maneja rutas con espacios y caracteres especiales correctamente.
"""
import os
import sys
import shutil
import subprocess

# Carpeta donde esta este script (= Aplicacion_Version_2\)
OUT_DIR = os.path.dirname(os.path.abspath(__file__))

# Carpeta del proyecto (un nivel arriba)
PROJECT_DIR = os.path.dirname(OUT_DIR)

SCRIPT_PY = os.path.join(PROJECT_DIR, "converter.py")
LOGO      = os.path.join(PROJECT_DIR, "IBC-Logo-2026_sin_fondo.png")
APP_NAME  = "Convertidor XML a Excel IBC"
TEMP_DIR  = os.path.join(OUT_DIR, "build_temp")

print(f"\n[INFO] Proyecto : {PROJECT_DIR}")
print(f"[INFO] Script   : {SCRIPT_PY}")
print(f"[INFO] Salida   : {OUT_DIR}")
print()

if not os.path.exists(SCRIPT_PY):
    print(f"ERROR: No se encontro converter.py en {PROJECT_DIR}")
    sys.exit(1)

# Argumentos para PyInstaller
args = [
    sys.executable, "-m", "PyInstaller",
    "--onefile",
    "--windowed",
    "--name", APP_NAME,
    "--distpath", OUT_DIR,
    "--workpath", TEMP_DIR,
    "--specpath", TEMP_DIR,
    "--clean",
]

if os.path.exists(LOGO):
    print(f"[INFO] Logo encontrado: {os.path.basename(LOGO)}")
    args += ["--add-data", f"{LOGO}{os.pathsep}."]
else:
    print("[INFO] Logo no encontrado, se compilara sin el.")

args.append(SCRIPT_PY)

print("[3/4] Compilando con PyInstaller...")
result = subprocess.run(args, cwd=PROJECT_DIR)

if result.returncode != 0:
    print("\nERROR: PyInstaller termino con error.")
    sys.exit(1)

# Limpiar temporales
if os.path.exists(TEMP_DIR):
    shutil.rmtree(TEMP_DIR, ignore_errors=True)
    print("[4/4] Temporales eliminados.")

exe_path = os.path.join(OUT_DIR, APP_NAME + ".exe")
if os.path.exists(exe_path):
    print("\n====================================================")
    print("  EXITO: Ejecutable generado en:")
    print(f"  {exe_path}")
    print("====================================================\n")
    os.startfile(OUT_DIR)
else:
    print("\nERROR: No se encontro el .exe generado.")
    sys.exit(1)
