import os
import sys
import subprocess

def install_pyinstaller():
    try:
        import PyInstaller  # noqa: F401
    except ImportError:
        print("Instalando PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])

def build():
    install_pyinstaller()
    main_script = "main.py"
    name = "rpa-utils"
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",
        "--noconsole",
        "--name", name,
        main_script
    ]
    print("Executando:", " ".join(cmd))
    subprocess.check_call(cmd)
    print("\nBuild finalizado! O executável estará em dist/")

if __name__ == "__main__":
    build()
