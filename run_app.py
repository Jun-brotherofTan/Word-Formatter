import os
import subprocess
import sys


def _run(cmd):
    return subprocess.run(cmd, check=False)


def _has_module(module_name):
    try:
        __import__(module_name)
        return True
    except Exception:
        return False


def ensure_pip_available():
    if _has_module("pip"):
        return True

    print("pip not found. Bootstrapping pip with ensurepip...")
    result = _run([sys.executable, "-m", "ensurepip", "--upgrade"])
    if result.returncode != 0:
        return False
    return _has_module("pip")


def install_dependencies():
    packages = ["python-docx", "requests"]

    if os.name == "nt":
        packages.append("pywin32")

    print("Installing dependencies:", ", ".join(packages))
    cmd = [sys.executable, "-m", "pip", "install", "--upgrade"] + packages
    result = _run(cmd)
    return result.returncode == 0


def main():
    if not ensure_pip_available():
        print("Failed to install pip automatically.")
        print("Install python with ensurepip support, then re-run this script.")
        sys.exit(1)

    if not install_dependencies():
        print("Failed to install required packages.")
        sys.exit(1)

    app_path = os.path.join(os.path.dirname(__file__), "wordFormatter.py")
    app_cmd = [sys.executable, app_path] + sys.argv[1:]
    sys.exit(subprocess.call(app_cmd))


if __name__ == "__main__":
    main()
