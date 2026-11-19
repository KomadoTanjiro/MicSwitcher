"""Профили: микрофон + вывод + пресеты эквалайзера под одним именем.

В config.json: "profiles": [{"name": "Гарнитура", "mic": "<id>", "out": "<id>",
"eq_mic": "Голос", "eq_out": "Бас", "hotkey": "ctrl+alt+1"}].
Пустой mic/out - устройство не меняется; eq_* отсутствует - пресет не меняется.
"""
from dataclasses import dataclass
from typing import Optional


@dataclass(frozen=True)
class Profile:
    name: str
    mic: str = ""
    out: str = ""
    eq_mic: Optional[str] = None
    eq_out: Optional[str] = None
    hotkey: str = ""

    @classmethod
    def from_dict(cls, data):
        if not isinstance(data, dict):
            return None
        name = str(data.get("name") or "").strip()
        if not name:
            return None

        def eq(key):
            value = data.get(key)
            return None if value is None else str(value)

        return cls(
            name=name,
            mic=str(data.get("mic") or "").strip(),
            out=str(data.get("out") or "").strip(),
            eq_mic=eq("eq_mic"),
            eq_out=eq("eq_out"),
            hotkey=str(data.get("hotkey") or "").strip(),
        )

    def to_dict(self):
        data = {"name": self.name, "mic": self.mic, "out": self.out}
        if self.eq_mic is not None:
            data["eq_mic"] = self.eq_mic
        if self.eq_out is not None:
            data["eq_out"] = self.eq_out
        data["hotkey"] = self.hotkey
        return data


def load(raw):
    """Профили из конфига: битые записи и повторы имён пропускаются."""
    result, seen = [], set()
    for item in raw if isinstance(raw, list) else []:
        profile = Profile.from_dict(item)
        if profile is None or profile.name.lower() in seen:
            continue
        seen.add(profile.name.lower())
        result.append(profile)
    return result


def dump(profiles):
    return [p.to_dict() for p in profiles]


def find(profiles, name):
    name = (name or "").strip().lower()
    return next((p for p in profiles if p.name.lower() == name), None)


def next_after(profiles, active_name):
    """Следующий профиль по кругу; если активного нет - первый."""
    if not profiles:
        return None
    current = find(profiles, active_name)
    if current is None:
        return profiles[0]
    return profiles[(profiles.index(current) + 1) % len(profiles)]


def upsert(profiles, profile):
    """Заменяет профиль с тем же именем (без учёта регистра) или добавляет в конец."""
    result, replaced = [], False
    for p in profiles:
        if p.name.lower() == profile.name.lower():
            result.append(profile)
            replaced = True
        else:
            result.append(p)
    if not replaced:
        result.append(profile)
    return result


def remove(profiles, name):
    return [p for p in profiles if p.name.lower() != (name or "").strip().lower()]


def auto_name(out_name, mic_name):
    parts = [n for n in (out_name, mic_name) if n]
    return " + ".join(parts) or "Профиль"
