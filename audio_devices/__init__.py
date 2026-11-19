"""Звуковые устройства Windows: общий интерфейс и выбор реализации.

Реализации (одинаковый набор методов):
- core.CoreAudioBackend - Core Audio напрямую через pycaw/comtypes, основная;
- svv.SvvBackend - внешний NirSoft SoundVolumeView.exe, запасная.

Методы реализации:
- list_devices(flow) - активные устройства направления, отсортированные;
- get_default_id(flow, role) - ID устройства по умолчанию для роли или None;
- set_default(device_id, roles) - сделать устройством по умолчанию для ролей.
Атрибуты: name - имя для лога, watchable - дешёвый ли опрос списка устройств.
"""
import os
from dataclasses import dataclass

CAPTURE = "capture"
RENDER = "render"

ROLE_CONSOLE = 0
ROLE_MULTIMEDIA = 1
ROLE_COMMUNICATIONS = 2
ALL_ROLES = (ROLE_CONSOLE, ROLE_MULTIMEDIA, ROLE_COMMUNICATIONS)

ROLE_ALIASES = {
    "0": (ROLE_CONSOLE,),
    "console": (ROLE_CONSOLE,),
    "1": (ROLE_MULTIMEDIA,),
    "multimedia": (ROLE_MULTIMEDIA,),
    "2": (ROLE_COMMUNICATIONS,),
    "communications": (ROLE_COMMUNICATIONS,),
    "all": ALL_ROLES,
}

BACKEND_MODES = ("auto", "core", "svv")


class AudioBackendError(RuntimeError):
    pass


def parse_roles(mode):
    """Значение set_default_mode ≥ роли Windows. Неизвестное = все роли."""
    return ROLE_ALIASES.get(str(mode).strip().lower(), ALL_ROLES)


@dataclass(frozen=True)
class Device:
    id: str
    name: str
    interface: str = ""
    flow: str = CAPTURE

    @property
    def label(self):
        if self.interface and self.interface != self.name:
            return f"{self.name} — {self.interface}"
        return self.name or self.id

    @property
    def legacy_id(self):
        """ID в формате SoundVolumeView (Command-Line Friendly ID): так устройства
        хранились в config.json до перехода на Core Audio."""
        suffix = "Capture" if self.flow == CAPTURE else "Render"
        return f"{self.interface}\\Device\\{self.name}\\{suffix}"


def sort_devices(devices):
    return sorted(devices, key=lambda d: (d.name.lower(), d.id.lower()))


def open_backend(mode, svv_path):
    """Возвращает (реализация, предупреждение или None).

    auto - Core Audio, при ошибке SoundVolumeView, если он лежит по svv_path;
    core - только Core Audio; svv - только SoundVolumeView."""
    mode = str(mode).strip().lower()

    if mode == "svv":
        from .svv import SvvBackend
        return SvvBackend(svv_path), None

    try:
        from .core import CoreAudioBackend
        return CoreAudioBackend(), None
    except Exception as e:
        if mode == "core" or not os.path.isfile(svv_path):
            raise AudioBackendError(f"Core Audio недоступен: {e}") from e
        from .svv import SvvBackend
        return SvvBackend(svv_path), f"Core Audio недоступен ({e}), работаю через SoundVolumeView."
