import os
import sys


def resource_path(relative_path: str) -> str:
    """Путь к встроенному ресурсу: из исходников — от корня проекта, из exe — из sys._MEIPASS."""
    if getattr(sys, "frozen", False):
        base = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, *relative_path.replace("/", os.sep).split(os.sep))

