@echo off
cd /d "%~dp0"
py -m PyInstaller --noconfirm --noconsole --onefile --name MicSwitcher --icon MicSwitcher.ico --hidden-import audio_devices.core --hidden-import audio_devices.svv mic_switcher_tray.py
