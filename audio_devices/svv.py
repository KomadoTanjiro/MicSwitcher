"""Запасная реализация через NirSoft SoundVolumeView.exe.

Каждый вызов - отдельный процесс; список устройств - CSV во временном файле.
"""
import csv
import io
import os
import subprocess
import tempfile
import threading
import time

from . import ALL_ROLES, CAPTURE, ROLE_COMMUNICATIONS, ROLE_CONSOLE, ROLE_MULTIMEDIA, Device, sort_devices

COL_NAME = "Name"
COL_ID = "Command-Line Friendly ID"
COLUMNS = (
    COL_NAME,
    "Direction",
    "Type",
    "Device State",
    "Default",
    "Default Communications",
    "Default Multimedia",
    COL_ID,
)
ROLE_COLUMNS = {
    ROLE_CONSOLE: "Default",
    ROLE_MULTIMEDIA: "Default Multimedia",
    ROLE_COMMUNICATIONS: "Default Communications",
}


def field(row, column):
    return (row.get(column) or "").strip()


def is_true(value):
    return str(value).strip().lower() in {"yes", "true", "1", "capture", "render"}


def safe_remove(path, retries=5, delay=0.15):
    for _ in range(retries):
        try:
            if os.path.exists(path):
                os.remove(path)
            return
        except OSError:
            time.sleep(delay)


def read_text_with_retry(path, attempts=8, delay=0.15):
    last_error = None
    for _ in range(attempts):
        try:
            if os.path.isfile(path):
                with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
                    return f.read()
        except OSError as e:
            last_error = e
        time.sleep(delay)
    raise RuntimeError(f"Не удалось прочитать временный CSV: {last_error}")


class SvvBackend:
    name = "SoundVolumeView"
    watchable = False

    def __init__(self, exe_path):
        if not os.path.isfile(exe_path):
            raise FileNotFoundError(f"Не найден файл SoundVolumeView.exe: {exe_path}")
        self.exe_path = exe_path
        self._lock = threading.Lock()

    def run(self, args):
        return subprocess.run(
            [self.exe_path] + args,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0,
        )

    def export_rows(self):
        with self._lock:
            fd, temp_csv = tempfile.mkstemp(prefix="mic_switcher_", suffix=".csv")
            os.close(fd)
            try:
                proc = self.run(["/scomma", temp_csv, "/Columns", ",".join(COLUMNS)])
                if proc.returncode != 0:
                    raise RuntimeError(
                        f"Не удалось получить список устройств. "
                        f"stdout={proc.stdout} stderr={proc.stderr}"
                    )
                data = read_text_with_retry(temp_csv)
                if not data.strip():
                    return []
                return list(csv.DictReader(io.StringIO(data)))
            finally:
                safe_remove(temp_csv)

    @staticmethod
    def _is_active(row, flow):
        direction = "capture" if flow == CAPTURE else "render"
        return (
            field(row, "Direction").lower() == direction
            and field(row, "Type").lower() == "device"
            and field(row, "Device State").lower() == "active"
        )

    def _active_rows(self, flow):
        return [r for r in self.export_rows() if self._is_active(r, flow)]

    def list_devices(self, flow):
        devices = []
        for row in self._active_rows(flow):
            dev_id = field(row, COL_ID)
            devices.append(
                Device(
                    id=dev_id,
                    name=field(row, COL_NAME) or dev_id,
                    interface=dev_id.split("\\")[0] if dev_id else "",
                    flow=flow,
                )
            )
        return sort_devices(devices)

    def get_default_id(self, flow, role):
        column = ROLE_COLUMNS[role]
        for row in self._active_rows(flow):
            if is_true(row.get(column, "")):
                return field(row, COL_ID)
        return None

    def set_default(self, device_id, roles):
        role_args = ["all"] if set(roles) == set(ALL_ROLES) else [str(r) for r in roles]
        for role in role_args:
            proc = self.run(["/SetDefault", device_id, role])
            if proc.returncode != 0:
                raise RuntimeError(f"SoundVolumeView: stdout={proc.stdout} stderr={proc.stderr}")
