"""Construiește aplicația din generator_oferte_bess.py:
    - pe Windows: 'Generator Oferte BESS.exe'
    - pe Mac:     'Generator Oferte BESS.app'
(PyInstaller face aplicația doar pentru sistemul pe care rulează: .app-ul se construiește pe un Mac.)

Rulare (din folderul în care e și generator_oferte_bess.py):
    Windows:  py -3.14 construieste_exe.py
    Mac:      python3 construieste_exe.py

Modulele interne din pandas / openpyxl / pptx / xlwings sunt listate direct din fișierele de pe disc
și date explicit lui PyInstaller, pentru că detecția lui automată nu le găsește pe toate.
"""
import importlib.util
import os
import sys

import PyInstaller
import PyInstaller.__main__

SCRIPT = "generator_oferte_bess.py"
NUME_APLICATIE = "Generator Oferte BESS"
PACHETE = ["pandas", "openpyxl", "pptx", "xlwings", "pypdf"]
PACHETE_MAC = ["appscript", "aem", "mactypes", "psutil"]      # xlwings pe Mac controlează Excel prin appscript
FOLDERE_EXCLUSE = {"tests", "__pycache__"}


def module_din_pachet(pachet):
    """Toate modulele (.py, .pyd, .so) din folderul pachetului, ex. 'pandas._libs._cyutility'."""
    spec = importlib.util.find_spec(pachet)
    if spec.submodule_search_locations is None:          # modul simplu (un singur fișier), nu pachet
        return [pachet]
    radacina = os.path.dirname(spec.origin)
    baza = os.path.dirname(radacina)
    module = set()
    for folder, subfoldere, fisiere in os.walk(radacina):
        subfoldere[:] = [s for s in subfoldere if s not in FOLDERE_EXCLUSE]
        for fisier in fisiere:
            if fisier.endswith((".py", ".pyd", ".so")):
                cale = os.path.relpath(os.path.join(folder, fisier), baza)
                nume = cale.split(".")[0].replace(os.sep, ".")      # 'interval.cp314-win_amd64.pyd' -> 'interval'
                if nume.endswith(".__init__"):
                    nume = nume[: -len(".__init__")]
                module.add(nume)
    return sorted(module)


def main():
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    mac = sys.platform == "darwin"
    print(f"Python {sys.version.split()[0]}, PyInstaller {PyInstaller.__version__}, sistem {sys.platform}")

    argumente = ["--noconfirm", "--clean", "--windowed", "--name", NUME_APLICATIE]
    if mac:
        argumente += ["--onedir", "--osx-bundle-identifier", "ro.taliva.generator-oferte-bess"]
        pachete = PACHETE + [p for p in PACHETE_MAC if importlib.util.find_spec(p)]
    else:
        argumente += ["--onefile", "--hidden-import", "win32timezone"]
        pachete = PACHETE
    for pachet in pachete:
        module = module_din_pachet(pachet)
        print(f"{pachet}: {len(module)} module incluse")
        if len(module) > 1:
            argumente += ["--collect-data", pachet, "--collect-binaries", pachet]
        for modul in module:
            argumente += ["--hidden-import", modul]
    for parte in sorted(f for f in os.listdir(".") if f.startswith("bess_") and f.endswith(".py")):
        argumente += ["--add-data", f"{parte}{os.pathsep}."]      # părțile programului, încărcate la pornire
    argumente.append(SCRIPT)

    PyInstaller.__main__.run(argumente)
    rezultat = f"{NUME_APLICATIE}.app" if mac else f"{NUME_APLICATIE}.exe"
    print(f'\nGata. Aplicatia e in folderul "dist": {os.path.join("dist", rezultat)}')


if __name__ == "__main__":
    main()
