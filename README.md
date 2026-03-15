# MicSwitcher

MicSwitcher is a small Windows tray utility for quickly switching between two microphones using a global hotkey.

## Features

- Automatically detects active microphone devices
- Lets you choose **Mic 1** and **Mic 2** from the tray menu
- Switches between selected microphones with a global hotkey
- Supports auto-start with Windows
- Creates `config.json` automatically on first launch

## Default hotkey

`Ctrl + Alt + M`

## Requirements

- Windows
- Python 3.11+ to run from source
- `SoundVolumeView.exe` placed next to the script or executable

## Install dependencies

```bash
py -m pip install -r requirements.txt