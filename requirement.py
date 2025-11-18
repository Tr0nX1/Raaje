import sys
import subprocess
import argparse

REQUIRED = [
    {"pip": "pandas", "import": "pandas"},
    {"pip": "openpyxl", "import": "openpyxl"},
    {"pip": "python-docx", "import": "docx"},
]

def check_package(pkg):
    name = pkg["import"]
    try:
        mod = __import__(name)
        ver = getattr(mod, "__version__", None)
        return True, ver
    except Exception:
        return False, None

def install_package(pip_name, upgrade=False):
    args = [sys.executable, "-m", "pip", "install"]
    if upgrade:
        args.append("--upgrade")
    args.append(pip_name)
    subprocess.check_call(args)

def write_requirements_txt(path="requirements.txt"):
    lines = []
    for pkg in REQUIRED:
        ok, ver = check_package(pkg)
        if ver:
            lines.append(f"{pkg['pip']}=={ver}")
        else:
            lines.append(pkg["pip"])
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")

def main():
    parser = argparse.ArgumentParser(description="Check and install required packages for Notice Generator")
    parser.add_argument("--install", action="store_true", help="Install missing packages")
    parser.add_argument("--upgrade", action="store_true", help="Upgrade required packages to latest")
    parser.add_argument("--write", action="store_true", help="Write requirements.txt with pinned versions")
    args = parser.parse_args()

    missing = []
    print("Checking packages:\n")
    for pkg in REQUIRED:
        ok, ver = check_package(pkg)
        label = pkg["pip"]
        if ok:
            print(f"  - {label}: OK{f' ({ver})' if ver else ''}")
        else:
            print(f"  - {label}: MISSING")
            missing.append(pkg)

    if args.install and missing:
        print("\nInstalling missing packages...\n")
        for pkg in missing:
            install_package(pkg["pip"], upgrade=args.upgrade)
        print("\nInstall complete")

    if args.upgrade and not args.install:
        print("\nUpgrading packages...\n")
        for pkg in REQUIRED:
            install_package(pkg["pip"], upgrade=True)
        print("\nUpgrade complete")

    if args.write:
        write_requirements_txt()
        print("\nWrote requirements.txt")

    if not args.install and missing:
        pkgs = " ".join(p["pip"] for p in missing)
        print(f"\nRun to install missing: {sys.executable} -m pip install {pkgs}")

if __name__ == "__main__":
    main()