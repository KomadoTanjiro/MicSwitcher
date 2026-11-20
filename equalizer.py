"""Эквалайзер через Equalizer APO.

Equalizer APO встраивается в звуковой тракт Windows и обрабатывает звук на уровне
устройства, поэтому эффект получают все программы (Discord, Steam, OBS, игры).
Он сам перечитывает файлы настроек при изменении.

MicSwitcher не трогает настройки пользователя: пишет свой файл MicSwitcher.txt в папку
config Equalizer APO и один раз дописывает в config.txt строку «Include: MicSwitcher.txt».
"""
import base64
import os
import re
import struct
import tempfile

INCLUDE_FILE = "MicSwitcher.txt"
INCLUDE_LINE = f"Include: {INCLUDE_FILE}"
OFF = "off"
OFF_LABEL = "Выкл"
DOWNLOAD_URL = "https://sourceforge.net/projects/equalizerapo/"
REG_PATH = r"SOFTWARE\EqualizerAPO"

# Шумодав - VST2-плагин RNNoise (werman/noise-suppression-for-voice).
# Equalizer APO работает внутри системного audiodg.exe, поэтому плагин должен лежать
# там, откуда его может прочитать служба: по умолчанию в <Equalizer APO>\VSTPlugins.
RNNOISE_URL = "https://github.com/werman/noise-suppression-for-voice/releases"
RNNOISE_TOKEN = "{rnnoise}"
RNNOISE_FILES = ("rnnoise_stereo.dll", "rnnoise_mono.dll")
RNNOISE_LINE = f'VSTPlugin: Library "{RNNOISE_TOKEN}"'

# Порог голоса (VAD) RNNoise, %: звук глушится, если вероятность голоса ниже порога.
# None - не передавать состояние, плагин берёт свои настройки (порог 60%).
RNNOISE_THRESHOLDS = (None, 0, 50, 70, 85, 90, 95)
RNNOISE_DEFAULT_THRESHOLD = 60
RNNOISE_RECOMMENDED = 90
RNNOISE_GRACE_UNITS = 20  # 10 мс за единицу, значение по умолчанию плагина
# Плагин собран на JUCE: в VST2 он работает в режиме «chunk», и Equalizer APO передаёт
# ему только ChunkData, а пары «имя значение» игнорирует. ChunkData - двоичное
# состояние JUCE (copyXmlToBinary): магическое число, длина XML, XML, нулевой байт.
JUCE_XML_MAGIC = 0x21324356

# Команды Equalizer APO. Preamp снижает уровень на величину подъёма, чтобы не было клиппинга.
BUILTIN_PRESETS = {
    "out": {
        "Бас": [
            "Preamp: -6 dB",
            "Filter: ON LSC Fc 105 Hz Gain 6 dB Q 0.7",
        ],
        "Чистый голос": [
            "Preamp: -3 dB",
            "Filter: ON HP Fc 90 Hz",
            "Filter: ON PK Fc 300 Hz Gain -2 dB Q 1.0",
            "Filter: ON PK Fc 3000 Hz Gain 3 dB Q 1.0",
        ],
        "Игры: шаги": [
            "Preamp: -5 dB",
            "Filter: ON LSC Fc 120 Hz Gain -3 dB Q 0.7",
            "Filter: ON PK Fc 2500 Hz Gain 5 dB Q 1.2",
            "Filter: ON HSC Fc 8000 Hz Gain 2 dB Q 0.7",
        ],
        "Мягкий": [
            "Filter: ON HSC Fc 6000 Hz Gain -4 dB Q 0.7",
        ],
    },
    "mic": {
        "Голос": [
            "Preamp: -3 dB",
            "Filter: ON HP Fc 80 Hz",
            "Filter: ON PK Fc 250 Hz Gain -3 dB Q 1.0",
            "Filter: ON PK Fc 3500 Hz Gain 3 dB Q 1.0",
        ],
        "Убрать гул": [
            "Filter: ON HP Fc 100 Hz",
            "Filter: ON NO Fc 50 Hz",
        ],
        "Шумодав": [
            RNNOISE_LINE,
        ],
        "Голос + шумодав": [
            RNNOISE_LINE,
            "Preamp: -3 dB",
            "Filter: ON HP Fc 80 Hz",
            "Filter: ON PK Fc 250 Hz Gain -3 dB Q 1.0",
            "Filter: ON PK Fc 3500 Hz Gain 3 dB Q 1.0",
        ],
        "Радио": [
            "Preamp: -4 dB",
            "Filter: ON HP Fc 100 Hz",
            "Filter: ON PK Fc 150 Hz Gain 2 dB Q 1.0",
            "Filter: ON PK Fc 4000 Hz Gain 4 dB Q 1.0",
        ],
    },
}

_ENDPOINT_GUID = re.compile(r"^\{[\d.]+\}\.\{([0-9a-fA-F-]{36})\}$")


def presets(target, custom=None):
    """Пресеты для target («mic» / «out»): встроенные, затем свои из конфига.

    Свои: {"Имя": {"target": "out", "lines": ["Filter: ..."]}}."""
    result = dict(BUILTIN_PRESETS.get(target, {}))
    for name, preset in (custom or {}).items():
        if not isinstance(preset, dict) or preset.get("target") != target:
            continue
        lines = [str(line) for line in preset.get("lines") or [] if str(line).strip()]
        if lines:
            result[str(name)] = lines
    return result


def preset_label(name):
    return OFF_LABEL if not name or name == OFF else name


def device_pattern(device_id, name):
    """Как указать устройство в команде Device: GUID конечной точки Core Audio без скобок,
    иначе (ID SoundVolumeView) - имя устройства.

    Device сравнивает каждое слово шаблона как подстроку с именем, подключением и GUID
    устройства, все слова должны найтись. GUID без скобок совпадёт при любом их написании."""
    match = _ENDPOINT_GUID.match(device_id or "")
    return match.group(1) if match else name


def render(sections):
    """sections: [(заголовок, подпись устройства, шаблон Device, имя пресета, команды)]."""
    lines = ["# MicSwitcher: файл пересоздаётся автоматически, ручные правки пропадут."]
    for title, device_label, pattern, preset, commands in sections:
        lines += ["", f"# {title}: {device_label}, пресет «{preset}»", f"Device: {pattern}"]
        lines += list(commands)
    lines += ["", "Device: all", ""]
    return "\n".join(lines)


def find_config_dir(override="", winreg=None, environ=None):
    """Папка config Equalizer APO или None, если он не установлен."""
    environ = os.environ if environ is None else environ
    candidates = []
    if override:
        candidates.append(override)
    else:
        if winreg is not None:
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, REG_PATH) as key:
                    candidates.append(winreg.QueryValueEx(key, "ConfigPath")[0])
            except OSError:
                pass
        program_files = environ.get("ProgramFiles")
        if program_files:
            candidates.append(os.path.join(program_files, "EqualizerAPO", "config"))
    for path in candidates:
        if path and os.path.isdir(path):
            return path
    return None


def has_include(text):
    pattern = re.compile(rf"^\s*Include\s*:\s*{re.escape(INCLUDE_FILE)}\s*$", re.IGNORECASE | re.MULTILINE)
    return bool(pattern.search(text))


def ensure_include(config_dir):
    """Дописывает Include в config.txt, если его там нет. True - если дописали."""
    path = os.path.join(config_dir, "config.txt")
    text = ""
    if os.path.isfile(path):
        with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
            text = f.read()
    if has_include(text):
        return False
    with open(path, "a", encoding="utf-8") as f:
        if text and not text.endswith("\n"):
            f.write("\n")
        f.write(INCLUDE_LINE + "\n")
    return True


def write_settings(config_dir, text):
    """MicSwitcher.txt через временный файл: Equalizer APO не увидит половину файла."""
    path = os.path.join(config_dir, INCLUDE_FILE)
    fd, tmp = tempfile.mkstemp(prefix=".micswitcher_", suffix=".txt", dir=config_dir)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(text)
        os.replace(tmp, path)
    except BaseException:
        try:
            os.remove(tmp)
        except OSError:
            pass
        raise


def settings_exist(config_dir):
    return os.path.isfile(os.path.join(config_dir, INCLUDE_FILE))


def install(config_dir, text):
    ensure_include(config_dir)
    write_settings(config_dir, text)


def needs_rnnoise(commands):
    return any(RNNOISE_TOKEN in line for line in commands)


def parse_threshold(value):
    """Порог из конфига: целое 0..100 или None (настройки плагина)."""
    if isinstance(value, bool) or value is None or value == "":
        return None
    try:
        number = int(value)
    except (TypeError, ValueError):
        return None
    return number if 0 <= number <= 100 else None


def threshold_label(value):
    if value is None:
        return f"По умолчанию плагина ({RNNOISE_DEFAULT_THRESHOLD}%)"
    if value == 0:
        return "0%: только шумоподавление"
    if value == RNNOISE_RECOMMENDED:
        return f"{value}% (рекомендуется)"
    return f"{value}%"


def rnnoise_chunk(threshold, grace_units=RNNOISE_GRACE_UNITS, retroactive_units=0):
    """Base64 состояния плагина RNNoise с нужным порогом для ChunkData."""
    xml = (
        '<RNNoise>'
        f'<PARAM id="vad_threshold" value="{threshold / 100:g}"/>'
        f'<PARAM id="vad_grace_period" value="{int(grace_units)}"/>'
        f'<PARAM id="vad_retroactive_grace_period" value="{int(retroactive_units)}"/>'
        '</RNNoise>'
    ).encode("utf-8")
    blob = struct.pack("<II", JUCE_XML_MAGIC, len(xml)) + xml + b"\0"
    return base64.b64encode(blob).decode("ascii")


def resolve_commands(commands, rnnoise_path, threshold=None):
    """Подставляет путь к плагину RNNoise и его порог.
    None - пресету нужен плагин, а его нет."""
    if needs_rnnoise(commands) and not rnnoise_path:
        return None
    chunk = f' ChunkData "{rnnoise_chunk(threshold)}"' if threshold is not None else ""
    result = []
    for line in commands:
        if RNNOISE_TOKEN in line:
            line = line.replace(RNNOISE_TOKEN, rnnoise_path) + chunk
        result.append(line)
    return result


def _is_vst2(path):
    """В релизе RNNoise dll с тем же именем есть и в .lv2 - это другой формат."""
    parts = os.path.normpath(path).lower().split(os.sep)
    return not any(p.endswith(".lv2") or p == "ladspa" for p in parts)


def _rnnoise_rank(path):
    """Стерео-версия лучше моно, x64 лучше x86, путь короче - лучше."""
    lower = path.lower()
    return (
        0 if lower.endswith(RNNOISE_FILES[0]) else 1,
        0 if "x64" in lower else 1,
        len(path),
    )


def find_rnnoise(config_dir, override="", max_depth=4):
    """Путь к rnnoise_*.dll или None. override из конфига, иначе поиск в
    <Equalizer APO>\VSTPlugins (папка рядом с config) на глубину max_depth."""
    if override:
        return override if os.path.isfile(override) else None
    if not config_dir:
        return None

    root = os.path.join(os.path.dirname(os.path.normpath(config_dir)), "VSTPlugins")
    found = []
    for folder, dirs, files in os.walk(root):
        depth = os.path.relpath(folder, root).count(os.sep)
        if depth >= max_depth:
            dirs[:] = []
        for name in files:
            path = os.path.join(folder, name)
            if name.lower() in RNNOISE_FILES and _is_vst2(path):
                found.append(path)
    return min(found, key=_rnnoise_rank) if found else None
