"""Автоустановка того, что нужно эквалайзеру: Equalizer APO и плагин шумодава RNNoise.

Версии и хэши закреплены: скачанный файл, который не совпал по SHA-256, не используется.

RNNoise ставится в <Equalizer APO>\\VSTPlugins\\RNNoise: Equalizer APO работает внутри
системного audiodg.exe и читает плагины оттуда. Источник - копия из релизного zip
MicSwitcher (plugins\\rnnoise), иначе архив с GitHub. Папка Equalizer APO лежит в
Program Files: если записать без прав не получилось, копирование повторяется
через PowerShell с запросом прав администратора (окно UAC).

Equalizer APO ставится его собственным установщиком: MicSwitcher скачивает и запускает
его, дальше пользователь отмечает устройства в Configurator и перезагружает компьютер.
"""
import base64
import hashlib
import os
import shutil
import subprocess
import tempfile
import urllib.request
import zipfile

USER_AGENT = "MicSwitcher"

RNNOISE_VERSION = "v1.10"
RNNOISE_ZIP_URL = (
    "https://github.com/werman/noise-suppression-for-voice/releases/download/v1.10/win-rnnoise.zip"
)
RNNOISE_ZIP_SHA256 = "e4dcc186102850f8dcebeb766ad1bf34c24f6280bd614ef84b72dae4783fc517"
# В архиве rnnoise_stereo.dll есть и в .lv2 (другой формат): берём только VST2 из vst/.
RNNOISE_ZIP_MEMBERS = {
    "rnnoise_stereo.dll": "win-rnnoise/vst/rnnoise_stereo.dll",
    "rnnoise_mono.dll": "win-rnnoise/vst/rnnoise_mono.dll",
}
RNNOISE_SOURCE_URL = "https://github.com/werman/noise-suppression-for-voice/tree/v1.10"
BUNDLED_RNNOISE_DIR = os.path.join("plugins", "rnnoise")
RNNOISE_TARGET_SUBDIR = "RNNoise"

EQAPO_VERSION = "1.4.2"
EQAPO_URL = "https://downloads.sourceforge.net/project/equalizerapo/1.4.2/EqualizerAPO-x64-1.4.2.exe"
EQAPO_SHA256 = "7403be7427bbe1936a40dded082829b6e217fc4f5990fee5cba501f0ae055afa"


class InstallError(RuntimeError):
    pass


def download(url, sha256, dest_path, opener=urllib.request.urlopen):
    """Скачивает url в dest_path и проверяет SHA-256. При несовпадении файл удаляется."""
    digest = hashlib.sha256()
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    try:
        with opener(request, timeout=120) as response, open(dest_path, "wb") as f:
            while True:
                chunk = response.read(1 << 16)
                if not chunk:
                    break
                digest.update(chunk)
                f.write(chunk)
    except OSError as e:
        raise InstallError(f"Не удалось скачать {url}: {e}") from e
    if digest.hexdigest() != sha256:
        os.remove(dest_path)
        raise InstallError(f"Скачанный файл не совпал по SHA-256: {url}")
    return dest_path


def bundled_rnnoise(app_dir):
    folder = os.path.join(app_dir, BUNDLED_RNNOISE_DIR)
    if all(os.path.isfile(os.path.join(folder, name)) for name in RNNOISE_ZIP_MEMBERS):
        return folder
    return None


def extract_rnnoise(zip_path, out_dir):
    os.makedirs(out_dir, exist_ok=True)
    with zipfile.ZipFile(zip_path) as archive:
        for name, member in RNNOISE_ZIP_MEMBERS.items():
            try:
                data = archive.read(member)
            except KeyError as e:
                raise InstallError(f"В архиве RNNoise нет {member}") from e
            with open(os.path.join(out_dir, name), "wb") as f:
                f.write(data)
    return out_dir


def rnnoise_target_dir(config_dir):
    return os.path.join(os.path.dirname(os.path.normpath(config_dir)), "VSTPlugins", RNNOISE_TARGET_SUBDIR)


def copy_files(src_dir, dest_dir, names):
    os.makedirs(dest_dir, exist_ok=True)
    for name in names:
        shutil.copy2(os.path.join(src_dir, name), os.path.join(dest_dir, name))


def _ps_quote(value):
    return "'" + str(value).replace("'", "''") + "'"


def elevated_copy_command(src_dir, dest_dir, names):
    """Команда: обычный PowerShell запускает PowerShell с правами администратора
    (-Verb RunAs, окно UAC) и ждёт его. Скрипт передаётся как -EncodedCommand."""
    sources = ",".join(_ps_quote(os.path.join(src_dir, n)) for n in names)
    script = (
        "$ErrorActionPreference = 'Stop'\n"
        f"New-Item -ItemType Directory -Force -Path {_ps_quote(dest_dir)} | Out-Null\n"
        f"Copy-Item -Force -LiteralPath {sources} -Destination {_ps_quote(dest_dir)}\n"
    )
    encoded = base64.b64encode(script.encode("utf-16-le")).decode("ascii")
    outer = (
        "Start-Process -FilePath powershell.exe -Verb RunAs -Wait -WindowStyle Hidden "
        f"-ArgumentList '-NoProfile','-NonInteractive','-EncodedCommand','{encoded}'"
    )
    return ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", outer]


def elevated_copy(src_dir, dest_dir, names, run=subprocess.run):
    run(
        elevated_copy_command(src_dir, dest_dir, names),
        capture_output=True,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    missing = [n for n in names if not os.path.isfile(os.path.join(dest_dir, n))]
    if missing:
        raise InstallError("Установка отменена или не удалась: нет прав на папку Equalizer APO.")


def install_rnnoise(app_dir, config_dir, run=subprocess.run, opener=urllib.request.urlopen):
    """Ставит RNNoise рядом с Equalizer APO. Возвращает путь к rnnoise_stereo.dll."""
    if not config_dir:
        raise InstallError("Сначала нужен Equalizer APO.")
    names = list(RNNOISE_ZIP_MEMBERS)
    dest_dir = rnnoise_target_dir(config_dir)

    with tempfile.TemporaryDirectory(prefix="micswitcher_rnnoise_") as tmp:
        src_dir = bundled_rnnoise(app_dir)
        if src_dir is None:
            zip_path = download(RNNOISE_ZIP_URL, RNNOISE_ZIP_SHA256, os.path.join(tmp, "win-rnnoise.zip"), opener)
            src_dir = extract_rnnoise(zip_path, os.path.join(tmp, "dll"))
        try:
            copy_files(src_dir, dest_dir, names)
        except PermissionError:
            elevated_copy(src_dir, dest_dir, names, run)

    return os.path.join(dest_dir, names[0])


def download_equalizer_apo(dest_dir=None, opener=urllib.request.urlopen):
    dest_dir = dest_dir or tempfile.mkdtemp(prefix="micswitcher_eqapo_")
    path = os.path.join(dest_dir, f"EqualizerAPO-x64-{EQAPO_VERSION}.exe")
    return download(EQAPO_URL, EQAPO_SHA256, path, opener)
