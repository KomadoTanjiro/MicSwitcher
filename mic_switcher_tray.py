import copy
import ctypes
import json
import os
import queue
import sys
import tempfile
import threading
import time
import tkinter as tk
import tkinter.simpledialog
from dataclasses import dataclass
from datetime import datetime

try:
    import winreg
except ImportError:  # не Windows: модуль импортируется только ради тестов
    winreg = None

import keyboard
import pystray
from PIL import Image, ImageDraw

import equalizer
import profiles as prof
from audio_devices import CAPTURE, RENDER, open_backend, parse_roles


APP_NAME = "MicSwitcher"
AUTORUN_REG_PATH = r"Software\Microsoft\Windows\CurrentVersion\Run"
SLOTS = (1, 2)
LOG_MAX_BYTES = 1_000_000
DEVICE_CACHE_TTL = 2.0
WATCH_INTERVAL = 2.0
TOGGLE_COOLDOWN = 0.3


@dataclass(frozen=True)
class Kind:
    """Вид устройств: микрофоны или вывод. Ключи конфига - selected_<key>_<N>, eq_<key>."""
    key: str
    flow: str
    short: str
    title: str
    noun: str
    plural: str
    hotkey_key: str


MIC = Kind("mic", CAPTURE, "Mic", "Микрофон", "микрофон", "микрофонов", "hotkey")
OUT = Kind("out", RENDER, "Out", "Вывод", "вывод", "устройств вывода", "hotkey_out")
KINDS = (MIC, OUT)

DEFAULT_CONFIG = {
    "audio_backend": "auto",
    "auto_refresh": True,
    "sound_volume_view_path": "SoundVolumeView.exe",
    "hotkey": "ctrl+alt+m",
    "hotkey_out": "ctrl+alt+o",
    "hotkey_profile": "ctrl+alt+p",
    "set_default_mode": "all",
    "log_file": "mic_switcher.log",
    "selected_mic_1": "",
    "selected_mic_2": "",
    "selected_out_1": "",
    "selected_out_2": "",
    "profiles": [],
    "active_profile": "",
    "eq_mic": equalizer.OFF,
    "eq_out": equalizer.OFF,
    "equalizer_apo_config_dir": "",
    "rnnoise_vst_path": "",
    "rnnoise_vad_threshold": None,
    "eq_custom_presets": {},
}


def get_app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


APP_DIR = get_app_dir()
CONFIG_PATH = os.path.join(APP_DIR, "config.json")

CONFIG = copy.deepcopy(DEFAULT_CONFIG)
SOUNDVOLUMEVIEW_EXE = ""
ROLES = parse_roles("all")
LOG_FILE = os.path.join(APP_DIR, DEFAULT_CONFIG["log_file"])

backend = None
icon = None
hotkey_handles = []
watch_stop = threading.Event()
_last_toggle = {}

switch_lock = threading.Lock()
config_lock = threading.Lock()
log_lock = threading.Lock()
cache_lock = threading.Lock()
_device_cache = {}


# --- Лог ---------------------------------------------------------------------

def log(message):
    line = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {message}"
    print(line)
    with log_lock:
        try:
            if os.path.isfile(LOG_FILE) and os.path.getsize(LOG_FILE) > LOG_MAX_BYTES:
                os.replace(LOG_FILE, LOG_FILE + ".1")
            with open(LOG_FILE, "a", encoding="utf-8") as f:
                f.write(line + "\n")
        except OSError:
            pass


# --- Файлы и конфиг ----------------------------------------------------------

def write_json_atomic(path, data):
    directory = os.path.dirname(path) or "."
    fd, tmp = tempfile.mkstemp(prefix=".tmp_", suffix=".json", dir=directory)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        os.replace(tmp, path)
    except BaseException:
        try:
            os.remove(tmp)
        except OSError:
            pass
        raise


def resolve_path(path_value):
    if os.path.isabs(path_value):
        return path_value
    return os.path.join(APP_DIR, path_value)


def read_config_file():
    """Возвращает (конфиг, предупреждение). Битый файл откладывается в .broken."""
    if not os.path.isfile(CONFIG_PATH):
        return {}, None

    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        if not isinstance(cfg, dict):
            raise ValueError("ожидался JSON-объект")
        return cfg, None
    except ValueError as e:
        backup = CONFIG_PATH + ".broken"
        try:
            os.replace(CONFIG_PATH, backup)
        except OSError:
            backup = "не сохранен"
        return {}, f"config.json поврежден ({e}), создан заново. Старый файл: {backup}"


def load_config():
    cfg, warning = read_config_file()
    merged = {**copy.deepcopy(DEFAULT_CONFIG), **cfg}
    if merged != cfg:
        write_json_atomic(CONFIG_PATH, merged)
    return merged, warning


def apply_config(cfg):
    global CONFIG, SOUNDVOLUMEVIEW_EXE, ROLES, LOG_FILE
    CONFIG = cfg
    SOUNDVOLUMEVIEW_EXE = resolve_path(str(cfg["sound_volume_view_path"]))
    ROLES = parse_roles(cfg["set_default_mode"])
    LOG_FILE = resolve_path(str(cfg["log_file"]))


def save_config():
    with config_lock:
        write_json_atomic(CONFIG_PATH, CONFIG)


# --- Уведомления и диалоги ---------------------------------------------------

class PopupManager:
    """Один поток tkinter на всё приложение: popup-ы и диалоги приходят через очередь."""

    BG = "#1f2937"
    WIDTH = 360
    MIN_HEIGHT = 90

    def __init__(self):
        self._queue = queue.Queue()
        self._thread = None
        self._lock = threading.Lock()

    def _ensure_thread(self):
        with self._lock:
            if self._thread is None or not self._thread.is_alive():
                self._thread = threading.Thread(target=self._run, daemon=True)
                self._thread.start()

    def show(self, title, message, duration_ms):
        self._ensure_thread()
        self._queue.put(("popup", title, message, duration_ms))

    def ask(self, title, prompt, initial, callback):
        """Окно ввода строки. callback(текст или None) вызывается в отдельном потоке."""
        self._ensure_thread()
        self._queue.put(("ask", title, prompt, initial, callback))

    def _run(self):
        try:
            root = tk.Tk()
            root.withdraw()
        except Exception as e:
            log(f"Не удалось показать popup: {e}")
            return

        current = {"popup": None}

        def poll():
            try:
                while True:
                    item = self._queue.get_nowait()
                    if item[0] == "ask":
                        self._ask(root, *item[1:])
                    else:
                        self._render(root, current, *item[1:])
            except queue.Empty:
                pass
            root.after(100, poll)

        root.after(0, poll)
        root.mainloop()

    @staticmethod
    def _ask(root, title, prompt, initial, callback):
        try:
            value = tkinter.simpledialog.askstring(title, prompt, initialvalue=initial, parent=root)
        except Exception as e:
            log(f"Не удалось показать окно ввода: {e}")
            value = None
        threading.Thread(target=callback, args=(value,), daemon=True).start()

    def _render(self, root, current, title, message, duration_ms):
        try:
            if current["popup"] is not None:
                current["popup"].destroy()

            popup = tk.Toplevel(root)
            current["popup"] = popup
            popup.overrideredirect(True)
            popup.attributes("-topmost", True)
            popup.configure(bg=self.BG)

            frame = tk.Frame(popup, bg=self.BG, bd=1, relief="solid")
            frame.pack(fill="both", expand=True)

            tk.Label(
                frame,
                text=title,
                font=("Segoe UI", 10, "bold"),
                bg=self.BG,
                fg="white",
                anchor="w",
            ).pack(fill="x", padx=12, pady=(10, 2))

            tk.Label(
                frame,
                text=message,
                font=("Segoe UI", 10),
                bg=self.BG,
                fg="#e5e7eb",
                justify="left",
                wraplength=self.WIDTH - 30,
                anchor="w",
            ).pack(fill="both", expand=True, padx=12, pady=(0, 10))

            popup.update_idletasks()
            height = max(self.MIN_HEIGHT, popup.winfo_reqheight())
            x = popup.winfo_screenwidth() - self.WIDTH - 20
            y = popup.winfo_screenheight() - height - 60
            popup.geometry(f"{self.WIDTH}x{height}+{x}+{y}")

            def close():
                if current["popup"] is popup:
                    current["popup"] = None
                popup.destroy()

            popup.after(duration_ms, close)
        except Exception as e:
            log(f"Не удалось показать popup: {e}")


popups = PopupManager()


def notify(title, message, timeout=3):
    log(f"{title}: {message}")
    popups.show(title, message, max(1500, int(timeout * 1000)))


def update_tray_title(text):
    log(text)
    if icon is not None:
        try:
            icon.title = text[:127]  # лимит подсказки трея в Windows
        except Exception as e:
            log(f"Не удалось обновить title tray icon: {e}")


def report(text, show_popup_flag=False, popup_title=APP_NAME, timeout=3):
    update_tray_title(text)
    if show_popup_flag:
        notify(popup_title, text, timeout=timeout)


def report_all(messages, timeout_ok=3, timeout_more=5):
    """Несколько сообщений одного действия - одним popup. Первое - главное."""
    lines = [m for m in messages if m]
    if lines:
        report("\n".join(lines), show_popup_flag=True, timeout=timeout_ok if len(lines) == 1 else timeout_more)


def show_fatal_error(message):
    log(f"Критическая ошибка: {message}")
    try:
        ctypes.windll.user32.MessageBoxW(None, message, APP_NAME, 0x10)
    except Exception:
        pass


# --- Звук --------------------------------------------------------------------

def open_audio():
    """Открывает реализацию по audio_backend. Возвращает предупреждение или None."""
    global backend
    backend, warning = open_backend(CONFIG["audio_backend"], SOUNDVOLUMEVIEW_EXE)
    invalidate_devices()
    log(f"Звук: {backend.name}")
    return warning


def invalidate_devices():
    with cache_lock:
        _device_cache.clear()


def get_devices(kind, force=False):
    """Активные устройства вида. Без force - кеш не старше DEVICE_CACHE_TTL."""
    with cache_lock:
        cached = _device_cache.get(kind.flow)
        if not force and cached is not None and time.monotonic() - cached[0] < DEVICE_CACHE_TTL:
            return list(cached[1])

    devices = backend.list_devices(kind.flow)

    with cache_lock:
        _device_cache[kind.flow] = (time.monotonic(), devices)
    return list(devices)


def get_current_id(kind):
    """ID текущего устройства для первой роли из set_default_mode (при all - Console)."""
    return backend.get_default_id(kind.flow, ROLES[0])


def find_device(kind, target_id):
    if not target_id:
        return None
    return next((d for d in get_devices(kind) if d.id == target_id), None)


def display_name(kind, target_id):
    device = find_device(kind, target_id)
    return device.name if device else "не подключен"


def set_default_device(kind, target_id):
    try:
        backend.set_default(target_id, ROLES)
    except Exception as e:
        raise RuntimeError(
            f"Не удалось переключить {kind.noun} на: {display_name(kind, target_id)}: {e}"
        ) from e


# --- Слоты: Mic 1 / Mic 2, Out 1 / Out 2 -------------------------------------

def slot_key(kind, slot):
    return f"selected_{kind.key}_{slot}"


def other_slot(slot):
    return 2 if slot == 1 else 1


def get_selected(kind, slot):
    return str(CONFIG.get(slot_key(kind, slot), "")).strip()


def selected_label(kind, slot):
    selected = get_selected(kind, slot)
    return display_name(kind, selected) if selected else "не выбран"


def migrate_legacy_selection(kind, devices):
    """ID SoundVolumeView из старых конфигов ≥ ID текущей реализации. Молча."""
    by_legacy = {d.legacy_id.lower(): d for d in devices}
    known = {d.id for d in devices}
    changed = False
    for slot in SLOTS:
        selected = get_selected(kind, slot)
        device = by_legacy.get(selected.lower())
        if selected and selected not in known and device is not None:
            CONFIG[slot_key(kind, slot)] = device.id
            log(f"{kind.short} {slot}: ID {selected} заменен на {device.id}")
            changed = True
    if changed:
        save_config()


def ensure_default_selection(kind, devices=None):
    """Дополняет выбор слотов из активных устройств.
    Возвращает текст для пользователя или None, если ничего не поменялось."""
    if devices is None:
        devices = get_devices(kind)
    migrate_legacy_selection(kind, devices)
    if len(devices) < 2:
        return f"Найдено меньше двух активных {kind.plural}."

    available = [d.id for d in devices]
    changed = False

    if get_selected(kind, 1) not in available:
        CONFIG[slot_key(kind, 1)] = available[0]
        changed = True

    first = get_selected(kind, 1)
    if get_selected(kind, 2) not in available or get_selected(kind, 2) == first:
        replacement = next((i for i in available if i != first), None)
        if replacement is not None:
            CONFIG[slot_key(kind, 2)] = replacement
            changed = True

    if not changed:
        return None

    save_config()
    return (
        f"{kind.title}: 1) {display_name(kind, get_selected(kind, 1))}, "
        f"2) {display_name(kind, get_selected(kind, 2))}"
    )


def ensure_all_selections():
    messages = []
    for kind in KINDS:
        try:
            messages.append(ensure_default_selection(kind))
        except Exception as e:
            messages.append(f"Ошибка списка устройств ({kind.noun}): {e}")
    return messages


def select_device(kind, slot, target_id):
    if target_id == get_selected(kind, other_slot(slot)):
        report(
            f"Нельзя выбрать одно и то же устройство и для {kind.short} 1, и для {kind.short} 2.",
            show_popup_flag=True,
            timeout=4,
        )
        return

    CONFIG[slot_key(kind, slot)] = target_id
    save_config()
    refresh_menu()
    report(f"{kind.short} {slot} выбран: {display_name(kind, target_id)}", show_popup_flag=True)


# --- Переключение ------------------------------------------------------------

def pick_toggle_target(current_id, first, second):
    """Текущий = слот 1 ≥ слот 2, во всех остальных случаях ≥ слот 1."""
    if current_id == first:
        return 2, second
    return 1, first


def switched(verb, kind, slot, target):
    """Сообщение о переключении + эквалайзер под новое устройство."""
    return [f"{verb} {kind.short} {slot}: {display_name(kind, target)}", apply_equalizer()]


def activate_slot(kind, slot):
    with switch_lock:
        try:
            target = get_selected(kind, slot)
            if not target:
                report(f"{kind.short} {slot} не выбран.", show_popup_flag=True)
                return

            set_default_device(kind, target)
            report_all(switched("Активен", kind, slot, target))
        except Exception as e:
            report(f"Ошибка {kind.short} {slot}: {e}", show_popup_flag=True, timeout=5)


def toggle(kind):
    if not switch_lock.acquire(blocking=False):
        log("Переключение уже выполняется, повторный вызов пропущен.")
        return

    try:
        now = time.monotonic()
        if now - _last_toggle.get(kind.key, float("-inf")) < TOGGLE_COOLDOWN:
            log("Повторное нажатие слишком быстро, пропущено.")
            return
        _last_toggle[kind.key] = now

        first, second = get_selected(kind, 1), get_selected(kind, 2)
        if not first or not second:
            report(
                f"Сначала выбери {kind.short} 1 и {kind.short} 2 в меню трея.",
                show_popup_flag=True,
                timeout=5,
            )
            return

        slot, target = pick_toggle_target(get_current_id(kind), first, second)
        set_default_device(kind, target)
        report_all(switched("Переключено на", kind, slot, target))
    except Exception as e:
        report(f"Ошибка переключения: {e}", show_popup_flag=True, timeout=5)
    finally:
        switch_lock.release()


def current_summary():
    lines = []
    for kind in KINDS:
        device = find_device(kind, get_current_id(kind))
        if device is None:
            lines.append(f"{kind.title}: не определен")
            continue
        slot = next((s for s in SLOTS if get_selected(kind, s) == device.id), None)
        prefix = f"{kind.short} {slot}, " if slot else ""
        lines.append(f"{kind.title}: {prefix}{device.label}")
    active = CONFIG.get("active_profile")
    if active:
        lines.append(f"Профиль: {active}")
    lines.append(
        f"EQ: вывод {equalizer.preset_label(CONFIG.get('eq_out'))}, "
        f"микрофон {equalizer.preset_label(CONFIG.get('eq_mic'))}"
    )
    return "\n".join(lines)


def show_current():
    try:
        report(current_summary(), show_popup_flag=True, timeout=5)
    except Exception as e:
        report(f"Ошибка определения текущих устройств: {e}", show_popup_flag=True, timeout=5)


def refresh_devices():
    try:
        messages = []
        for kind in KINDS:
            devices = get_devices(kind, force=True)
            messages.append(f"{kind.title}: найдено {len(devices)}")
            messages.append(ensure_default_selection(kind, devices))
        refresh_menu()
        report_all(["Список устройств обновлен."] + messages)
    except Exception as e:
        report(f"Ошибка обновления списка устройств: {e}", show_popup_flag=True, timeout=5)


# --- Эквалайзер --------------------------------------------------------------

def eq_key(kind):
    return f"eq_{kind.key}"


def eq_presets(kind):
    return equalizer.presets(kind.key, CONFIG.get("eq_custom_presets"))


def eq_config_dir():
    return equalizer.find_config_dir(str(CONFIG.get("equalizer_apo_config_dir") or ""), winreg)


def rnnoise_path(config_dir=None):
    config_dir = config_dir or eq_config_dir()
    return equalizer.find_rnnoise(config_dir, str(CONFIG.get("rnnoise_vst_path") or ""))


def rnnoise_threshold():
    return equalizer.parse_threshold(CONFIG.get("rnnoise_vad_threshold"))


def mic_uses_rnnoise():
    return equalizer.needs_rnnoise(eq_presets(MIC).get(CONFIG.get("eq_mic"), []))


def set_rnnoise_threshold(value):
    CONFIG["rnnoise_vad_threshold"] = value
    save_config()
    refresh_menu()
    messages = [f"Порог шумодава: {equalizer.threshold_label(value)}"]
    if not mic_uses_rnnoise():
        messages.append("Применится, когда у микрофона пресет «Шумодав» или «Голос + шумодав».")
    messages.append(apply_equalizer())
    report_all(messages)


def eq_active():
    return any(CONFIG.get(eq_key(k), equalizer.OFF) != equalizer.OFF for k in KINDS)


def apply_equalizer():
    """Пишет MicSwitcher.txt под текущие устройства по умолчанию.
    Возвращает текст проблемы или None."""
    config_dir = eq_config_dir()
    if config_dir is None:
        return "Equalizer APO не найден: эквалайзер не применен." if eq_active() else None
    if not eq_active() and not equalizer.settings_exist(config_dir):
        return None

    sections, problems = [], []
    for kind in KINDS:
        preset = CONFIG.get(eq_key(kind), equalizer.OFF)
        commands = eq_presets(kind).get(preset)
        if preset == equalizer.OFF or not commands:
            continue
        device = find_device(kind, get_current_id(kind))
        if device is None:
            continue
        commands = equalizer.resolve_commands(commands, rnnoise_path(config_dir), rnnoise_threshold())
        if commands is None:
            problems.append(
                f"Шумодав: плагин RNNoise не найден. Распакуй rnnoise_stereo.dll в папку "
                f"VSTPlugins Equalizer APO или укажи rnnoise_vst_path в config.json."
            )
            continue
        pattern = equalizer.device_pattern(device.id, device.name)
        sections.append((kind.title, device.label, pattern, preset, commands))

    try:
        equalizer.install(config_dir, equalizer.render(sections))
    except OSError as e:
        problems.append(f"Не удалось записать настройки Equalizer APO: {e}")
    return "\n".join(problems) or None


def set_eq(kind, preset):
    CONFIG[eq_key(kind)] = preset
    save_config()
    refresh_menu()
    report_all([f"Эквалайзер ({kind.noun}): {equalizer.preset_label(preset)}", apply_equalizer()])


def open_target(target):
    try:
        os.startfile(target)
    except Exception as e:
        report(f"Не удалось открыть {target}: {e}", show_popup_flag=True, timeout=5)


# --- Профили -----------------------------------------------------------------

def get_profiles():
    return prof.load(CONFIG.get("profiles"))


def set_profiles(items):
    CONFIG["profiles"] = prof.dump(items)
    save_config()


def apply_profile(name):
    with switch_lock:
        profile = prof.find(get_profiles(), name)
        if profile is None:
            report(f"Профиль «{name}» не найден.", show_popup_flag=True, timeout=4)
            return

        messages = [f"Профиль «{profile.name}»"]
        for kind, target in ((OUT, profile.out), (MIC, profile.mic)):
            if not target:
                continue
            if find_device(kind, target) is None:
                messages.append(f"{kind.title} не подключен")
                continue
            try:
                set_default_device(kind, target)
                messages.append(f"{kind.title}: {display_name(kind, target)}")
            except RuntimeError as e:
                messages.append(str(e))

        for kind in KINDS:
            preset = getattr(profile, eq_key(kind))
            if preset is not None:
                CONFIG[eq_key(kind)] = preset
        CONFIG["active_profile"] = profile.name
        save_config()

        messages.append(apply_equalizer())
        refresh_menu()
        report_all(messages)


def next_profile():
    profile = prof.next_after(get_profiles(), CONFIG.get("active_profile"))
    if profile is None:
        report("Профилей нет: меню «Профиль», пункт «Сохранить текущее как профиль».", show_popup_flag=True, timeout=5)
        return
    apply_profile(profile.name)


def build_current_profile(name, existing=None):
    return prof.Profile(
        name=name,
        mic=get_current_id(MIC) or "",
        out=get_current_id(OUT) or "",
        eq_mic=CONFIG.get("eq_mic", equalizer.OFF),
        eq_out=CONFIG.get("eq_out", equalizer.OFF),
        hotkey=existing.hotkey if existing else "",
    )


def save_profile_named(name):
    name = (name or "").strip()
    if not name:
        return
    items = get_profiles()
    existing = prof.find(items, name)
    set_profiles(prof.upsert(items, build_current_profile(name, existing)))
    CONFIG["active_profile"] = name
    save_config()
    errors = register_hotkeys()
    refresh_menu()
    action = "обновлен" if existing else "сохранен"
    report_all([f"Профиль «{name}» {action}.", *errors])


def save_current_profile():
    try:
        out = find_device(OUT, get_current_id(OUT))
        mic = find_device(MIC, get_current_id(MIC))
        initial = prof.auto_name(out.name if out else "", mic.name if mic else "")
    except Exception as e:
        report(f"Ошибка определения текущих устройств: {e}", show_popup_flag=True, timeout=5)
        return
    popups.ask(APP_NAME, "Название профиля:", initial, save_profile_named)


def delete_profile(name):
    set_profiles(prof.remove(get_profiles(), name))
    if CONFIG.get("active_profile", "").lower() == name.lower():
        CONFIG["active_profile"] = ""
        save_config()
    register_hotkeys()
    refresh_menu()
    report(f"Профиль «{name}» удален.", show_popup_flag=True)


# --- Подключение и отключение устройств --------------------------------------

def diff_devices(kind, previous, current):
    """Строки о подключённых и отключённых устройствах. Аргументы - {id: Device}."""
    if previous is None or previous.keys() == current.keys():
        return []
    added = sorted(current[i].label for i in current.keys() - previous.keys())
    removed = sorted(previous[i].label for i in previous.keys() - current.keys())
    return [f"Подключен {kind.noun}: {n}" for n in added] + [f"Отключен {kind.noun}: {n}" for n in removed]


def watch_devices(stop):
    """Опрос устройств: меню и popup при подключении и отключении, эквалайзер
    за устройством по умолчанию, даже если его сменили в настройках Windows.
    Выбор слотов не трогаем: выдернутая гарнитура вернётся на своё место."""
    previous = {kind.key: None for kind in KINDS}
    defaults = None
    while not stop.wait(WATCH_INTERVAL):
        try:
            lines = []
            for kind in KINDS:
                current = {d.id: d for d in get_devices(kind, force=True)}
                lines += diff_devices(kind, previous[kind.key], current)
                previous[kind.key] = current

            current_defaults = tuple(get_current_id(kind) for kind in KINDS)
            eq_problem = None
            if defaults is not None and current_defaults != defaults:
                eq_problem = apply_equalizer()
            defaults = current_defaults

            if lines:
                refresh_menu()
            if lines or eq_problem:
                report_all(lines + [eq_problem], timeout_ok=4)
        except Exception as e:
            log(f"Ошибка опроса устройств: {e}")


def start_watcher():
    global watch_stop
    if not CONFIG.get("auto_refresh") or not backend.watchable:
        return
    watch_stop = threading.Event()
    threading.Thread(target=watch_devices, args=(watch_stop,), daemon=True).start()
    log("Опрос устройств запущен.")


def stop_watcher():
    watch_stop.set()


# --- Хоткеи ------------------------------------------------------------------

def hotkey_bindings():
    """(сочетание, функция, аргументы, описание). Пустое сочетание - хоткей выключен."""
    bindings = [(CONFIG.get(kind.hotkey_key), toggle, (kind,), f"переключить {kind.noun}") for kind in KINDS]
    bindings.append((CONFIG.get("hotkey_profile"), next_profile, (), "следующий профиль"))
    for profile in get_profiles():
        bindings.append((profile.hotkey, apply_profile, (profile.name,), f"профиль «{profile.name}»"))
    return [(str(c).strip(), fn, args, title) for c, fn, args, title in bindings if c and str(c).strip()]


def register_hotkeys():
    """keyboard слушает клавиатуру в своём потоке. Возвращает список ошибок."""
    for handle in hotkey_handles:
        try:
            keyboard.remove_hotkey(handle)
        except (KeyError, ValueError) as e:
            log(f"Ошибка снятия hotkey: {e}")
    hotkey_handles.clear()

    errors = []
    for combo, fn, args, title in hotkey_bindings():
        try:
            hotkey_handles.append(keyboard.add_hotkey(combo, fn, args=args))
            log(f"Hotkey {combo}: {title}")
        except Exception as e:
            errors.append(f"Не удалось зарегистрировать hotkey {combo} ({title}): {e}")
    return errors


def reload_config():
    try:
        cfg, warning = load_config()
        apply_config(cfg)
        stop_watcher()

        messages = ["Конфиг перечитан.", warning, *register_hotkeys()]
        try:
            messages.append(open_audio())
            messages += ensure_all_selections()
            messages.append(apply_equalizer())
            start_watcher()
        except Exception as e:
            messages.append(str(e))

        refresh_menu()
        report_all(messages)
    except Exception as e:
        report(f"Не удалось перечитать config: {e}", show_popup_flag=True, timeout=5)


# --- Автозапуск --------------------------------------------------------------

def get_gui_python():
    """pythonw.exe вместо python.exe, чтобы при автозапуске не было консоли."""
    folder, name = os.path.split(sys.executable)
    if name.lower() == "python.exe":
        candidate = os.path.join(folder, "pythonw.exe")
        if os.path.isfile(candidate):
            return candidate
    return sys.executable


def get_autorun_command():
    if getattr(sys, "frozen", False):
        return f'"{sys.executable}"'
    return f'"{get_gui_python()}" "{os.path.abspath(__file__)}"'


def is_autorun_enabled():
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, AUTORUN_REG_PATH, 0, winreg.KEY_READ) as key:
            value, _ = winreg.QueryValueEx(key, APP_NAME)
            return value == get_autorun_command()
    except FileNotFoundError:
        return False
    except OSError as e:
        log(f"Ошибка чтения автозапуска: {e}")
        return False


def set_autorun(enabled):
    action = "включить" if enabled else "выключить"
    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER, AUTORUN_REG_PATH, 0, winreg.KEY_SET_VALUE
        ) as key:
            if enabled:
                winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, get_autorun_command())
            else:
                try:
                    winreg.DeleteValue(key, APP_NAME)
                except FileNotFoundError:
                    pass

        refresh_menu()
        report(f"Автозапуск {'включен' if enabled else 'выключен'}.", show_popup_flag=True)
    except OSError as e:
        report(f"Не удалось {action} автозапуск: {e}", show_popup_flag=True, timeout=5)


def toggle_autorun():
    set_autorun(not is_autorun_enabled())


# --- Файлы в редакторе -------------------------------------------------------

def open_config():
    try:
        os.startfile(CONFIG_PATH)
    except Exception as e:
        report(f"Не удалось открыть config: {e}", show_popup_flag=True, timeout=5)


def open_log():
    try:
        if not os.path.exists(LOG_FILE):
            with open(LOG_FILE, "a", encoding="utf-8"):
                pass
        os.startfile(LOG_FILE)
    except Exception as e:
        report(f"Не удалось открыть лог: {e}", show_popup_flag=True, timeout=5)


# --- Трей --------------------------------------------------------------------

def create_image():
    image = Image.new("RGB", (64, 64), (30, 30, 30))
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle((8, 8, 56, 56), radius=10, fill=(60, 120, 220))
    draw.rectangle((28, 18, 36, 36), fill=(255, 255, 255))
    draw.ellipse((22, 32, 42, 50), fill=(255, 255, 255))
    draw.rectangle((30, 46, 34, 56), fill=(255, 255, 255))
    return image


def menu_action(fn, *args):
    def handler(icon_obj, item):
        fn(*args)
    return handler


def make_checked(getter, value):
    def checker(item):
        return getter() == value
    return checker


def disabled_item(text):
    return pystray.MenuItem(text, menu_action(lambda: None), enabled=False)


def hotkey_hint(combo):
    combo = str(combo or "").strip()
    return f"  ({combo})" if combo else ""


def build_select_menu(kind, slot):
    devices = get_devices(kind)
    if not devices:
        return pystray.Menu(disabled_item("Нет устройств"))
    return pystray.Menu(
        *[
            pystray.MenuItem(
                d.label,
                menu_action(select_device, kind, slot, d.id),
                checked=make_checked(lambda s=slot: get_selected(kind, s), d.id),
                radio=True,
            )
            for d in devices
        ]
    )


def build_kind_menu(kind):
    return pystray.Menu(
        *[
            pystray.MenuItem(f"Включить {kind.short} {s}: {selected_label(kind, s)}", menu_action(activate_slot, kind, s))
            for s in SLOTS
        ],
        pystray.Menu.SEPARATOR,
        *[pystray.MenuItem(f"Выбрать {kind.short} {s}", build_select_menu(kind, s)) for s in SLOTS],
    )


def build_profiles_menu():
    items = get_profiles()
    active = lambda: CONFIG.get("active_profile", "")  # noqa: E731
    entries = [
        pystray.MenuItem(
            f"{p.name}{hotkey_hint(p.hotkey)}",
            menu_action(apply_profile, p.name),
            checked=make_checked(active, p.name),
            radio=True,
        )
        for p in items
    ] or [disabled_item("Профилей пока нет")]
    entries += [
        pystray.Menu.SEPARATOR,
        pystray.MenuItem(f"Следующий профиль{hotkey_hint(CONFIG.get('hotkey_profile'))}", menu_action(next_profile)),
        pystray.MenuItem("Сохранить текущее как профиль...", menu_action(save_current_profile)),
    ]
    if items:
        entries.append(
            pystray.MenuItem(
                "Удалить профиль",
                pystray.Menu(*[pystray.MenuItem(p.name, menu_action(delete_profile, p.name)) for p in items]),
            )
        )
    return pystray.Menu(*entries)


def build_eq_kind_menu(kind):
    current = lambda: CONFIG.get(eq_key(kind), equalizer.OFF)  # noqa: E731
    names = [equalizer.OFF] + list(eq_presets(kind))
    return pystray.Menu(
        *[
            pystray.MenuItem(
                equalizer.preset_label(name),
                menu_action(set_eq, kind, name),
                checked=make_checked(current, name),
                radio=True,
            )
            for name in names
        ]
    )


def build_threshold_menu():
    return pystray.Menu(
        *[
            pystray.MenuItem(
                equalizer.threshold_label(value),
                menu_action(set_rnnoise_threshold, value),
                checked=make_checked(rnnoise_threshold, value),
                radio=True,
            )
            for value in equalizer.RNNOISE_THRESHOLDS
        ]
    )


def build_eq_menu():
    config_dir = eq_config_dir()
    items = [pystray.MenuItem(kind.title, build_eq_kind_menu(kind)) for kind in (OUT, MIC)]
    items.append(pystray.MenuItem("Порог шумодава", build_threshold_menu()))
    items.append(pystray.Menu.SEPARATOR)
    if config_dir is None:
        items.append(pystray.MenuItem("Скачать Equalizer APO", menu_action(open_target, equalizer.DOWNLOAD_URL)))
    else:
        items.append(pystray.MenuItem("Открыть папку Equalizer APO", menu_action(open_target, config_dir)))
        if rnnoise_path(config_dir) is None:
            items.append(pystray.MenuItem("Скачать шумодав RNNoise", menu_action(open_target, equalizer.RNNOISE_URL)))
    return pystray.Menu(*items)


def build_menu():
    active = CONFIG.get("active_profile") or "нет"
    return pystray.Menu(
        *[
            pystray.MenuItem(f"Переключить {kind.noun}{hotkey_hint(CONFIG.get(kind.hotkey_key))}", menu_action(toggle, kind))
            for kind in KINDS
        ],
        pystray.MenuItem(f"Профиль: {active}", build_profiles_menu()),
        pystray.Menu.SEPARATOR,
        *[pystray.MenuItem(kind.title, build_kind_menu(kind)) for kind in KINDS],
        pystray.MenuItem("Эквалайзер", build_eq_menu()),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Показать текущие", menu_action(show_current)),
        pystray.MenuItem("Обновить список устройств", menu_action(refresh_devices)),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem(
            "Автозапуск",
            menu_action(toggle_autorun),
            checked=lambda item: is_autorun_enabled(),
        ),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Открыть config.json", menu_action(open_config)),
        pystray.MenuItem("Перечитать config.json", menu_action(reload_config)),
        pystray.MenuItem("Открыть лог", menu_action(open_log)),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Выход", menu_action(quit_app)),
    )


def refresh_menu():
    if icon is not None:
        try:
            icon.menu = build_menu()
            icon.update_menu()
        except Exception as e:
            log(f"Не удалось обновить меню: {e}")


def quit_app():
    log("Завершение работы.")
    stop_watcher()
    try:
        keyboard.unhook_all_hotkeys()
    except Exception as e:
        log(f"Ошибка unhook hotkeys: {e}")

    if icon is not None:
        try:
            icon.stop()
        except Exception as e:
            log(f"Ошибка остановки icon: {e}")


# --- Запуск ------------------------------------------------------------------

def main():
    global icon

    cfg, warning = load_config()
    apply_config(cfg)
    log("Старт приложения.")

    try:
        backend_warning = open_audio()
    except Exception as e:
        show_fatal_error(
            f"Не удалось подключиться к звуковым устройствам Windows.\n\n{e}\n\n"
            f"Переустановите зависимости (py -m pip install -r requirements.txt) "
            f"или положите SoundVolumeView.exe рядом с MicSwitcher."
        )
        return 1

    messages = ["Приложение запущено.", warning, backend_warning]
    messages += ensure_all_selections()
    messages += register_hotkeys()
    messages.append(apply_equalizer())

    icon = pystray.Icon(APP_NAME, create_image(), APP_NAME, menu=build_menu())
    report_all(messages, timeout_ok=2)
    start_watcher()

    log("Tray icon создан, вызываю icon.run()")
    icon.run()
    log("icon.run() завершился")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as e:
        log(f"Критическая ошибка: {e}")
        raise
