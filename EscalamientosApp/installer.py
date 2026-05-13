"""
EscalamientosApp - Instalador
"""
import os, sys, shutil, tempfile, subprocess

APP_NAME = "EscalamientosApp"
DESTINO = os.path.join(os.environ["USERPROFILE"], APP_NAME)

def crear_acceso_directo():
    try:
        bat_path = os.path.join(DESTINO, "iniciar.bat")
        ps_script = (
            '$desktop = [Environment]::GetFolderPath("Desktop"); '
            '$sc_path = Join-Path $desktop "EscalamientosApp.lnk"; '
            '$ws = New-Object -ComObject WScript.Shell; '
            '$sc = $ws.CreateShortcut($sc_path); '
            '$sc.TargetPath = "' + bat_path + '"; '
            '$sc.WorkingDirectory = "' + DESTINO + '"; '
            '$sc.Description = "Gestion de Escalamientos ATM - BHD"; '
            '$sc.Save()'
        )
        ps_file = os.path.join(tempfile.gettempdir(), "_installer_shortcut.ps1")
        with open(ps_file, 'w', encoding='utf-8') as f:
            f.write(ps_script)
        subprocess.run(['powershell', '-ExecutionPolicy', 'Bypass', '-File', ps_file], capture_output=True, creationflags=0x08000000)
        try: os.remove(ps_file)
        except: pass
    except: pass

def instalar():
    print("Instalando EscalamientosApp...")
    print(f"Destino: {DESTINO}")
    xlsx_dest = os.path.join(DESTINO, "backend", "PlanillaEscalamientos.xlsx")
    xlsx_backup = None
    if os.path.exists(xlsx_dest):
        xlsx_backup = xlsx_dest + ".backup"
        shutil.copy2(xlsx_dest, xlsx_backup)
    if os.path.exists(DESTINO):
        shutil.rmtree(DESTINO, ignore_errors=True)
    os.makedirs(DESTINO, exist_ok=True)
    for item in ['backend', 'iniciar.bat', 'detener_servidor.bat']:
        src = os.path.join(sys._MEIPASS, item)
        dst = os.path.join(DESTINO, item)
        if os.path.isdir(src):
            shutil.copytree(src, dst, ignore=shutil.ignore_patterns('__pycache__', '*.pyc'))
        elif os.path.exists(src):
            shutil.copy2(src, dst)
    if xlsx_backup and os.path.exists(xlsx_backup):
        if os.path.exists(xlsx_dest): os.remove(xlsx_dest)
        shutil.move(xlsx_backup, xlsx_dest)
    venv_path = os.path.join(DESTINO, "backend", "venv")
    if os.path.exists(venv_path):
        shutil.rmtree(venv_path, ignore_errors=True)
    crear_acceso_directo()
    return True

if __name__ == '__main__':
    instalar()
    subprocess.run(['powershell', '-Command',
        '$r = (New-Object -ComObject WScript.Shell).Popup("EscalamientosApp instalado. Usa el acceso directo en tu escritorio.\n\nQueres iniciar la aplicacion ahora?", 0, "Instalacion completada", 4); '
        'if ($r -eq 6) { Start-Process "' + os.path.join(DESTINO, "iniciar.bat") + '" -WorkingDirectory "' + DESTINO + '" }'],
        creationflags=0x08000000)
