"""Заглушки Windows-only библиотек: тесты логики идут на любой ОС."""
import os
import sys
import types

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)


def install():
    for name in ("keyboard", "pystray"):
        sys.modules.setdefault(name, types.ModuleType(name))
    try:
        import PIL  # noqa: F401
    except ImportError:
        pil = types.ModuleType("PIL")
        pil.Image = types.ModuleType("PIL.Image")
        pil.ImageDraw = types.ModuleType("PIL.ImageDraw")
        sys.modules.update({"PIL": pil, "PIL.Image": pil.Image, "PIL.ImageDraw": pil.ImageDraw})


install()
