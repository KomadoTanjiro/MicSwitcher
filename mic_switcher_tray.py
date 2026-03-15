import csv
import io
import json
import os
import subprocess
import sys
import tempfile
import threading
import time
import winreg
import tkinter as tk
from datetime import datetime

import keyboard
import pystray
from PIL import Image, ImageDraw


APP_NAME = "MicSwitcher"
AUTORUN_REG_PATH = r"Software\Microsoft\Windows\CurrentVersion\Run"


def get_app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


APP_DIR = get_app_dir()
CONFIG_PATH = os.path.join(APP_DIR, "config.json")
DEFAULT_CONFIG = {
    "sound_volume_view_path": "SoundVolumeView.exe",
    "hotkey": "ctrl+alt+m",
    "set_default_mode": "all",
    "log_file": "mic_switcher.log",
    "selected_mic_1": "",
    "selected_mic_2": "",
}

icon = None
switch_lock = threading.Lock()
file_lock = threading.Lock()
config_lock = threading.Lock()


def ensure_config_exists():
    if os.path.isfile(CONFIG_PATH):
        return

    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(DEFAULT_CONFIG, f, ensure_ascii=False, indent=2)


def load_config():
    ensure_config_exists()

    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        cfg = json.load(f)

    changed = False
    for key, value in DEFAULT_CONFIG.items():
        if key not in cfg:
            cfg[key] = value
            changed = True

    if changed:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)

    return cfg


def save_config():
    with config_lock:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(CONFIG, f, ensure_ascii=False, indent=2)


def resolve_path(path_value):
    if os.path.isabs(path_value):
        return path_value
    return os.path.join(APP_DIR, path_value)


CONFIG = load_config()
SOUNDVOLUMEVIEW_EXE = resolve_path(CONFIG["sound_volume_view_path"])
HOTKEY = CONFIG["hotkey"]
SET_DEFAULT_MODE = CONFIG["set_default_mode"]
LOG_FILE = resolve_path(CONFIG["log_file"])


def log(message):
    line = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {message}"
    print(line)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def show_popup(title, message, duration_ms=2500):
    def worker():
        try:
            root = tk.Tk()
            root.withdraw()

            popup = tk.Toplevel(root)
            popup.overrideredirect(True)
            popup.attributes("-topmost", True)
            popup.configure(bg="#1f2937")

            width = 360
            height = 90

            screen_width = popup.winfo_screenwidth()
            screen_height = popup.winfo_screenheight()

            x = screen_width - width - 20
            y = screen_height - height - 60

            popup.geometry(f"{width}x{height}+{x}+{y}")

            frame = tk.Frame(popup, bg="#1f2937", bd=1, relief="solid")
            frame.pack(fill="both", expand=True)

            title_label = tk.Label(
                frame,
                text=title,
                font=("Segoe UI", 10, "bold"),
                bg="#1f2937",
                fg="white",
                anchor="w",
            )
            title_label.pack(fill="x", padx=12, pady=(10, 2))

            msg_label = tk.Label(
                frame,
                text=message,
                font=("Segoe UI", 10),
                bg="#1f2937",
                fg="#e5e7eb",
                justify="left",
                wraplength=330,
                anchor="w",
            )
            msg_label.pack(fill="both", expand=True, padx=12, pady=(0, 10))

            popup.after(duration_ms, root.destroy)
            root.mainloop()
        except Exception as e:
            log(f"Не удалось показать popup: {e}")

    threading.Thread(target=worker, daemon=True).start()


def notify(title, message, timeout=3):
    log(f"{title}: {message}")
    duration_ms = max(1500, int(timeout * 1000))
    show_popup(title, message, duration_ms=duration_ms)


def ensure_tool_exists():
    if not os.path.isfile(SOUNDVOLUMEVIEW_EXE):
        raise FileNotFoundError(
            f"Не найден файл SoundVolumeView.exe: {SOUNDVOLUMEVIEW_EXE}"
        )


def run_svv(args):
    cmd = [SOUNDVOLUMEVIEW_EXE] + args
    return subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0,
    )


def safe_remove(path, retries=5, delay=0.15):
    for _ in range(retries):
        try:
            if os.path.exists(path):
                os.remove(path)
            return
        except PermissionError:
            time.sleep(delay)
        except OSError:
            time.sleep(delay)


def export_devices():
    columns = (
        "Name,Direction,Type,Device State,Default,Default Communications,"
        "Default Multimedia,Command-Line Friendly ID"
    )

    with file_lock:
        fd, temp_csv = tempfile.mkstemp(prefix="mic_switcher_", suffix=".csv")
        os.close(fd)

        try:
            proc = run_svv(["/scomma", temp_csv, "/Columns", columns])

            if proc.returncode != 0:
                raise RuntimeError(
                    f"Не удалось получить список устройств. "
                    f"stdout={proc.stdout} stderr={proc.stderr}"
                )

            data = None
            last_error = None

            for _ in range(8):
                try:
                    if os.path.isfile(temp_csv):
                        with open(temp_csv, "r", encoding="utf-8-sig", errors="replace") as f:
                            data = f.read()
                        break
                except PermissionError as e:
                    last_error = e
                    time.sleep(0.15)
                except OSError as e:
                    last_error = e
                    time.sleep(0.15)

            if data is None:
                raise RuntimeError(f"Не удалось прочитать временный CSV: {last_error}")

            if not data.strip():
                return []

            return list(csv.DictReader(io.StringIO(data)))

        finally:
            safe_remove(temp_csv)


def get_capture_devices():
    devices = export_devices()
    result = []

    for d in devices:
        direction = (d.get("Direction") or "").strip().lower()
        item_type = (d.get("Type") or "").strip().lower()
        state = (d.get("Device State") or "").strip().lower()

        if direction == "capture" and item_type == "device" and state == "active":
            result.append(d)

    result.sort(
        key=lambda x: (
            (x.get("Name") or "").lower(),
            (x.get("Command-Line Friendly ID") or "").lower(),
        )
    )
    return result


def is_true(value):
    return str(value).strip().lower() in {"yes", "true", "1", "capture"}


def get_current_default_capture():
    devices = get_capture_devices()
    for d in devices:
        if (
            is_true(d.get("Default", ""))
            or is_true(d.get("Default Multimedia", ""))
            or is_true(d.get("Default Communications", ""))
        ):
            return d
    return None


def get_selected_mic_1():
    return CONFIG.get("selected_mic_1", "").strip()


def get_selected_mic_2():
    return CONFIG.get("selected_mic_2", "").strip()


def set_selected_mic_1(device_id):
    CONFIG["selected_mic_1"] = device_id
    save_config()


def set_selected_mic_2(device_id):
    CONFIG["selected_mic_2"] = device_id
    save_config()


def get_device_by_id(device_id):
    if not device_id:
        return None

    devices = get_capture_devices()
    for d in devices:
        current_id = (d.get("Command-Line Friendly ID") or "").strip()
        if current_id == device_id:
            return d
    return None


def get_display_name(device_id):
    device = get_device_by_id(device_id)
    if device:
        return device.get("Name", device_id)
    return device_id


def get_friendly_label(device):
    name = (device.get("Name") or "").strip()
    device_id = (device.get("Command-Line Friendly ID") or "").strip()
    vendor_part = device_id.split("\\")[0] if device_id else ""

    if vendor_part and vendor_part != name:
        return f"{name} — {vendor_part}"
    return name or device_id


def set_default_mic(device_name_or_id):
    proc = run_svv(["/SetDefault", device_name_or_id, SET_DEFAULT_MODE])
    if proc.returncode != 0:
        raise RuntimeError(
            f"Не удалось переключить микрофон на: {device_name_or_id}\n"
            f"stdout={proc.stdout}\nstderr={proc.stderr}"
        )


def update_tray_title(text):
    global icon
    log(text)
    if icon is not None:
        try:
            icon.title = text
        except Exception as e:
            log(f"Не удалось обновить title tray icon: {e}")


def report(text, show_popup_flag=False, popup_title=APP_NAME, timeout=3):
    update_tray_title(text)
    if show_popup_flag:
        notify(popup_title, text, timeout=timeout)


def ensure_default_selection():
    devices = get_capture_devices()
    if len(devices) < 2:
        report("Найдено меньше двух активных микрофонов.", show_popup_flag=True, timeout=5)
        return

    mic_1 = get_selected_mic_1()
    mic_2 = get_selected_mic_2()
    available_ids = [(d.get("Command-Line Friendly ID") or "").strip() for d in devices]

    changed = False

    if mic_1 not in available_ids:
        CONFIG["selected_mic_1"] = available_ids[0]
        changed = True

    if mic_2 not in available_ids or CONFIG["selected_mic_2"] == CONFIG["selected_mic_1"]:
        for device_id in available_ids:
            if device_id != CONFIG["selected_mic_1"]:
                CONFIG["selected_mic_2"] = device_id
                changed = True
                break

    if changed:
        save_config()
        report(
            f"Выбраны микрофоны: 1) {get_display_name(CONFIG['selected_mic_1'])}, "
            f"2) {get_display_name(CONFIG['selected_mic_2'])}",
            show_popup_flag=True,
            timeout=4,
        )


def get_executable_for_autorun():
    return sys.executable


def get_script_for_autorun():
    if getattr(sys, "frozen", False):
        return None
    return os.path.abspath(__file__)


def get_autorun_command():
    exe = get_executable_for_autorun()

    if getattr(sys, "frozen", False):
        return f'"{exe}"'

    script_path = get_script_for_autorun()
    return f'"{exe}" "{script_path}"'


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


def enable_autorun():
    try:
        command = get_autorun_command()
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            AUTORUN_REG_PATH,
            0,
            winreg.KEY_SET_VALUE,
        ) as key:
            winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, command)

        refresh_menu()
        report("Автозапуск включен.", show_popup_flag=True)
    except OSError as e:
        report(f"Не удалось включить автозапуск: {e}", show_popup_flag=True, timeout=5)


def disable_autorun():
    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            AUTORUN_REG_PATH,
            0,
            winreg.KEY_SET_VALUE,
        ) as key:
            try:
                winreg.DeleteValue(key, APP_NAME)
            except FileNotFoundError:
                pass

        refresh_menu()
        report("Автозапуск выключен.", show_popup_flag=True)
    except OSError as e:
        report(f"Не удалось выключить автозапуск: {e}", show_popup_flag=True, timeout=5)


def toggle_autorun():
    if is_autorun_enabled():
        disable_autorun()
    else:
        enable_autorun()


def refresh_menu():
    global icon
    if icon is not None:
        try:
            icon.menu = build_menu()
            icon.update_menu()
        except Exception as e:
            log(f"Не удалось обновить меню: {e}")


def select_mic_1(device_id):
    selected_mic_2 = get_selected_mic_2()
    if device_id == selected_mic_2:
        report(
            "Нельзя выбрать один и тот же микрофон и для Mic 1, и для Mic 2.",
            show_popup_flag=True,
            timeout=4,
        )
        return

    set_selected_mic_1(device_id)
    refresh_menu()
    report(f"Mic 1 выбран: {get_display_name(device_id)}", show_popup_flag=True)


def select_mic_2(device_id):
    selected_mic_1 = get_selected_mic_1()
    if device_id == selected_mic_1:
        report(
            "Нельзя выбрать один и тот же микрофон и для Mic 1, и для Mic 2.",
            show_popup_flag=True,
            timeout=4,
        )
        return

    set_selected_mic_2(device_id)
    refresh_menu()
    report(f"Mic 2 выбран: {get_display_name(device_id)}", show_popup_flag=True)


def set_mic_1_now():
    with switch_lock:
        try:
            mic_1 = get_selected_mic_1()
            if not mic_1:
                report("Mic 1 не выбран.", show_popup_flag=True)
                return

            set_default_mic(mic_1)
            time.sleep(0.2)
            report(f"Активен Mic 1: {get_display_name(mic_1)}", show_popup_flag=True)
        except Exception as e:
            report(f"Ошибка mic_1: {e}", show_popup_flag=True, timeout=5)


def set_mic_2_now():
    with switch_lock:
        try:
            mic_2 = get_selected_mic_2()
            if not mic_2:
                report("Mic 2 не выбран.", show_popup_flag=True)
                return

            set_default_mic(mic_2)
            time.sleep(0.2)
            report(f"Активен Mic 2: {get_display_name(mic_2)}", show_popup_flag=True)
        except Exception as e:
            report(f"Ошибка mic_2: {e}", show_popup_flag=True, timeout=5)


def toggle_mic():
    if not switch_lock.acquire(blocking=False):
        log("Переключение уже выполняется, повторный вызов пропущен.")
        return

    try:
        mic_1 = get_selected_mic_1()
        mic_2 = get_selected_mic_2()

        if not mic_1 or not mic_2:
            report("Сначала выбери Mic 1 и Mic 2 в меню трея.", show_popup_flag=True, timeout=5)
            return

        current = get_current_default_capture()
        current_id = (current or {}).get("Command-Line Friendly ID", "").strip()

        if current_id == mic_1:
            target = mic_2
            label = "Mic 2"
        else:
            target = mic_1
            label = "Mic 1"

        set_default_mic(target)
        time.sleep(0.2)
        report(f"Переключено на {label}: {get_display_name(target)}", show_popup_flag=True)

    except Exception as e:
        report(f"Ошибка переключения: {e}", show_popup_flag=True, timeout=5)

    finally:
        switch_lock.release()


def show_current_mic():
    try:
        current = get_current_default_capture()
        if not current:
            report("Текущий микрофон не определен.", show_popup_flag=True, timeout=4)
            return

        name = current.get("Name", "")
        current_id = current.get("Command-Line Friendly ID", "")
        report(f"Текущий микрофон: {name} | {current_id}", show_popup_flag=True, timeout=4)
    except Exception as e:
        report(f"Ошибка определения текущего микрофона: {e}", show_popup_flag=True, timeout=5)


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


def refresh_devices():
    try:
        devices = get_capture_devices()
        if not devices:
            report("Активные микрофоны не найдены.", show_popup_flag=True, timeout=4)
            return

        ensure_default_selection()
        refresh_menu()
        report(f"Список микрофонов обновлен. Найдено: {len(devices)}", show_popup_flag=True)
    except Exception as e:
        report(f"Ошибка обновления списка устройств: {e}", show_popup_flag=True, timeout=5)


def create_image():
    image = Image.new("RGB", (64, 64), (30, 30, 30))
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle((8, 8, 56, 56), radius=10, fill=(60, 120, 220))
    draw.rectangle((28, 18, 36, 36), fill=(255, 255, 255))
    draw.ellipse((22, 32, 42, 50), fill=(255, 255, 255))
    draw.rectangle((30, 46, 34, 56), fill=(255, 255, 255))
    return image


def quit_app(icon_obj=None, item=None):
    global icon
    log("Завершение работы.")
    try:
        keyboard.unhook_all_hotkeys()
    except Exception as e:
        log(f"Ошибка unhook hotkeys: {e}")

    if icon is not None:
        try:
            icon.stop()
        except Exception as e:
            log(f"Ошибка остановки icon: {e}")


def hotkey_worker():
    try:
        keyboard.add_hotkey(HOTKEY, toggle_mic)
        log(f"Hotkey зарегистрирован: {HOTKEY}")
        keyboard.wait()
    except Exception as e:
        log(f"Ошибка hotkey_worker: {e}")


def make_select_mic_1_handler(device_id):
    def handler(icon_obj, item):
        select_mic_1(device_id)
    return handler


def make_select_mic_2_handler(device_id):
    def handler(icon_obj, item):
        select_mic_2(device_id)
    return handler


def make_checked_mic_1(device_id):
    def checker(item):
        return get_selected_mic_1() == device_id
    return checker


def make_checked_mic_2(device_id):
    def checker(item):
        return get_selected_mic_2() == device_id
    return checker


def build_select_mic_1_menu():
    devices = get_capture_devices()

    if not devices:
        return pystray.Menu(
            pystray.MenuItem("Нет устройств", lambda icon, item: None, enabled=False)
        )

    items = []
    for device in devices:
        device_id = (device.get("Command-Line Friendly ID") or "").strip()
        label = get_friendly_label(device)
        items.append(
            pystray.MenuItem(
                label,
                make_select_mic_1_handler(device_id),
                checked=make_checked_mic_1(device_id),
                radio=True,
            )
        )
    return pystray.Menu(*items)


def build_select_mic_2_menu():
    devices = get_capture_devices()

    if not devices:
        return pystray.Menu(
            pystray.MenuItem("Нет устройств", lambda icon, item: None, enabled=False)
        )

    items = []
    for device in devices:
        device_id = (device.get("Command-Line Friendly ID") or "").strip()
        label = get_friendly_label(device)
        items.append(
            pystray.MenuItem(
                label,
                make_select_mic_2_handler(device_id),
                checked=make_checked_mic_2(device_id),
                radio=True,
            )
        )
    return pystray.Menu(*items)


def on_toggle(icon_obj, item):
    toggle_mic()


def on_set_mic_1_now(icon_obj, item):
    set_mic_1_now()


def on_set_mic_2_now(icon_obj, item):
    set_mic_2_now()


def on_show_current(icon_obj, item):
    show_current_mic()


def on_open_config(icon_obj, item):
    open_config()


def on_open_log(icon_obj, item):
    open_log()


def on_refresh(icon_obj, item):
    refresh_devices()


def on_toggle_autorun(icon_obj, item):
    toggle_autorun()


def on_exit(icon_obj, item):
    quit_app(icon_obj, item)


def build_menu():
    mic_1_name = get_display_name(get_selected_mic_1()) if get_selected_mic_1() else "не выбран"
    mic_2_name = get_display_name(get_selected_mic_2()) if get_selected_mic_2() else "не выбран"

    return pystray.Menu(
        pystray.MenuItem("Переключить микрофон", on_toggle),
        pystray.MenuItem(f"Включить Mic 1: {mic_1_name}", on_set_mic_1_now),
        pystray.MenuItem(f"Включить Mic 2: {mic_2_name}", on_set_mic_2_now),
        pystray.MenuItem("Показать текущий", on_show_current),
        pystray.MenuItem("Обновить список устройств", on_refresh),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Выбрать Mic 1", build_select_mic_1_menu()),
        pystray.MenuItem("Выбрать Mic 2", build_select_mic_2_menu()),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem(
            "Автозапуск",
            on_toggle_autorun,
            checked=lambda item: is_autorun_enabled(),
        ),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Открыть config.json", on_open_config),
        pystray.MenuItem("Открыть лог", on_open_log),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Выход", on_exit),
    )


def main():
    global icon

    log("Старт приложения.")
    ensure_tool_exists()
    log("SoundVolumeView найден.")

    ensure_default_selection()

    hotkey_thread = threading.Thread(target=hotkey_worker, daemon=True)
    hotkey_thread.start()
    log("Поток hotkey запущен.")

    image = create_image()
    log("Иконка создана.")

    icon = pystray.Icon(APP_NAME, image, APP_NAME)
    icon.menu = build_menu()

    report("Приложение запущено.", show_popup_flag=True, timeout=2)
    log("Tray icon создан, вызываю icon.run()")
    icon.run()
    log("icon.run() завершился")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log(f"Критическая ошибка: {e}")
        raise