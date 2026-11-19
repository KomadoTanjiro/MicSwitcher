# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

MicSwitcher - утилита для Windows в трее: хоткеи переключают микрофон (`ctrl+alt+m`, Mic 1/2) и вывод (`ctrl+alt+o`, Out 1/2) по умолчанию, профили (`ctrl+alt+p`) меняют микрофон + вывод + эквалайзер сразу, эквалайзер - через Equalizer APO. Python 3.11+, звук напрямую через Windows Core Audio (pycaw/comtypes), NirSoft SoundVolumeView - только запасной вариант. Подробности: `docs/01-product.md` … `docs/05-tasks.md`. `docs/03-development.md` - обязательные правила, `docs/05-tasks.md` - план и статусы (обновлять при завершении задач).

## Команды

```bat
py -m pip install -r requirements.txt
pyw mic_switcher_tray.py
build.bat
py -m unittest -v
py -m unittest tests.test_app.ProfileTests.test_apply_profile
```

С Mac/Linux: `python3 -m unittest -v` из корня (системный Python 3.9 macOS подходит). Тесты подменяют `keyboard`, `pystray`, Pillow (`tests/stubs.py`), реализацию звука (`FakeBackend` в `tests/test_app.py`), `comtypes`/`pycaw` (`tests/test_core_audio.py`), запуск SoundVolumeView (`FakeRun`); папка Equalizer APO - временная (`equalizer_apo_config_dir`). Трей, хоткеи, реестр, popup, окно ввода, настоящий COM и Equalizer APO тестами не покрыты: перед merge ручная проверка на Windows по `docs/03-development.md`. Линтера нет. Код и тесты держать совместимыми с Python 3.9.

## Архитектура

- **Слои.** `mic_switcher_tray.py` - приложение (трей, меню, хоткеи, слоты, профили, опрос, popup, автозапуск). `audio_devices/` - звук: `Device(id, name, interface, flow)`, роли, `open_backend(mode, svv_path)`; реализация обязана уметь `list_devices(flow)`, `get_default_id(flow, role)`, `set_default(id, roles)` и иметь `name`, `watchable`. `profiles.py` и `equalizer.py` - чистая логика без глобального состояния. В приложении никаких COM-вызовов, SoundVolumeView и прямой записи в папку Equalizer APO.
- **Виды и слоты.** `Kind` (`MIC` = `CAPTURE`, `OUT` = `RENDER`) + слот `1`/`2`: `get_selected(kind, slot)`, `select_device`, `activate_slot`, `toggle(kind)`, `get_devices(kind)`; ключи конфига `selected_<mic|out>_<N>`, `eq_<mic|out>`. Копий функций под вид не заводить.
- **Core Audio (`core.py`).** `IMMDeviceEnumerator` для списка и текущего, `IPolicyConfig.SetDefaultEndpoint` для переключения (недокументирован). `CoInitialize` один раз на поток через `threading.local`, COM-объекты создаются на каждый вызов. ID устройства вида `{0.0.1.00000000}.{guid}`.
- **SoundVolumeView (`svv.py`).** CSV через `/scomma`, ID = `Command-Line Friendly ID`. `audio_backend: auto` уходит на него, только если Core Audio не создался и exe лежит рядом. Старые ID в конфиге переводятся через `Device.legacy_id` (`migrate_legacy_selection`).
- **Профили.** `Profile(name, mic, out, eq_mic, eq_out, hotkey)`: пустой `mic`/`out` и `eq_* = None` - не трогать. `apply_profile` под `switch_lock`; сохранение - `popups.ask` (simpledialog в потоке `PopupManager`), ответ в `save_profile_named`.
- **Эквалайзер.** `equalizer.find_config_dir` (конфиг, `HKLM\SOFTWARE\EqualizerAPO\ConfigPath`, `%ProgramFiles%`), `apply_equalizer()` пишет `MicSwitcher.txt` с секциями `Device: <guid без скобок>` + команды пресета для текущих устройств (шумодав - `VSTPlugin: Library "<rnnoise_stereo.dll>"`, путь ищет `find_rnnoise` в `<Equalizer APO>\\VSTPlugins` или берёт `rnnoise_vst_path`; порог `rnnoise_vad_threshold` - только через `ChunkData` в формате JUCE `copyXmlToBinary`, пары параметров плагин в режиме chunk игнорирует) и один раз добавляет `Include: MicSwitcher.txt` в `config.txt`. Вызывается после любого переключения (`switched`), профиля, выбора пресета и смены устройства по умолчанию вне приложения (`watch_devices`).
- **Хоткеи.** `hotkey_bindings()` (`hotkey`, `hotkey_out`, `hotkey_profile`, хоткеи профилей; пустые пропускаются), `register_hotkeys()` снимает старые и регистрирует заново, ошибки возвращает списком.
- **Потоки.** Главный - pystray и обработчики меню (`menu_action`). Хоткеи - поток `keyboard`. `watch_devices` - опрос раз в 2 с, только при `auto_refresh` и `backend.watchable`; меню и popup, выбор слотов не трогает. Popup и окно ввода - один поток `PopupManager`; tkinter трогать только там. `switch_lock` + `TOGGLE_COOLDOWN` 0.3 с на вид.
- **Конфиг.** Импорт модулей без побочных эффектов: `main()` - `load_config()`, `apply_config()`, `open_audio()`. `DEFAULT_CONFIG` копируется `deepcopy`, профили меняются через `set_profiles`. Битый JSON - `config.json.broken`, запись атомарная. `DEFAULT_CONFIG` и `config.example.json` одинаковые (есть тест).
- **Обратная связь.** `report` - подсказка иконки, лог (ротация 1 МБ), popup; `report_all` - несколько сообщений одного действия одним popup; вспомогательные функции возвращают текст. После изменения выбора, профилей, пресетов, конфига или устройств - `refresh_menu()`.
- **Автозапуск.** `HKCU\Software\Microsoft\Windows\CurrentVersion\Run\MicSwitcher`, из исходников через `pythonw.exe`; включён, только если значение совпадает с `get_autorun_command()`.

## Git

- `main`, фичи - отдельными ветками через PR в `KomadoTanjiro/MicSwitcher`.
- Автор - как в истории: `KomadoTanjiro <gorshok.09@gmail.com>`. Без `Co-Authored-By` и подписей генераторов.
- Тексты интерфейса, лога и документации - на русском.
