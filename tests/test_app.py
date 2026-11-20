import json
import os
import sys
import tempfile
import unittest
from unittest import mock

from tests import stubs

import audio_devices as ad
import mic_switcher_tray as ms
from mic_switcher_tray import MIC, OUT

HEADSET = ad.Device("{headset}", "Headset", "Jabra")
DESK = ad.Device("{desk}", "Desk Mic", "USB Audio")
CAMERA = ad.Device("{camera}", "Camera", "Logi")
PHONES = ad.Device("{0.0.0.00000000}.{11111111-1111-1111-1111-111111111111}", "Наушники", "Jabra", ad.RENDER)
SPEAKERS = ad.Device("{0.0.0.00000000}.{22222222-2222-2222-2222-222222222222}", "Динамики", "Realtek", ad.RENDER)


def installer_error(text):
    return ms.installer.InstallError(text)


class FakeBackend:
    name = "Fake"
    watchable = True

    def __init__(self, devices):
        self.devices = list(devices)
        self.defaults = {}
        self.set_calls = []
        self.list_calls = 0
        self.fail = False

    def list_devices(self, flow):
        self.list_calls += 1
        return ad.sort_devices(d for d in self.devices if d.flow == flow)

    def get_default_id(self, flow, role):
        return self.defaults.get((flow, role))

    def set_default(self, device_id, roles):
        if self.fail:
            raise OSError("boom")
        self.set_calls.append((device_id, tuple(roles)))
        flow = next(d.flow for d in self.devices if d.id == device_id)
        for role in roles:
            self.defaults[(flow, role)] = device_id


class Base(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = self.tmp.name
        self.eq_dir = os.path.join(self.dir, "eqapo")
        self.clock = [1000.0]
        patches = [
            mock.patch.object(ms, "APP_DIR", self.dir),
            mock.patch.object(ms, "CONFIG_PATH", os.path.join(self.dir, "config.json")),
            mock.patch.object(ms, "popups", mock.Mock()),
            mock.patch.object(ms, "icon", None),
            mock.patch.object(ms, "_last_toggle", {}),
            mock.patch.object(ms, "hotkey_handles", []),
            mock.patch.object(ms, "keyboard", mock.Mock()),
            mock.patch("time.monotonic", side_effect=lambda: self.clock[0]),
            mock.patch("builtins.print"),
        ]
        for p in patches:
            p.start()
            self.addCleanup(p.stop)
        self.addCleanup(self.tmp.cleanup)
        cfg, _ = ms.load_config()
        self.eq_dir = os.path.join(self.dir, "EqualizerAPO", "config")
        cfg["equalizer_apo_config_dir"] = self.eq_dir
        ms.apply_config(cfg)
        self.backend = self.use([HEADSET, DESK, PHONES, SPEAKERS])

    def use(self, devices):
        backend = FakeBackend(devices)
        p = mock.patch.object(ms, "backend", backend)
        p.start()
        self.addCleanup(p.stop)
        ms.invalidate_devices()
        return backend

    def tick(self, seconds=1.0):
        self.clock[0] += seconds

    def install_eq(self):
        os.makedirs(self.eq_dir, exist_ok=True)

    def eq_text(self):
        with open(os.path.join(self.eq_dir, "MicSwitcher.txt"), encoding="utf-8") as f:
            return f.read()

    def saved_config(self):
        with open(ms.CONFIG_PATH, encoding="utf-8") as f:
            return json.load(f)

    def popup_texts(self):
        return [c.args[1] for c in ms.popups.show.call_args_list]

    def select_all(self):
        ms.CONFIG.update(
            selected_mic_1="{headset}",
            selected_mic_2="{desk}",
            selected_out_1=PHONES.id,
            selected_out_2=SPEAKERS.id,
        )


class ConfigTests(Base):
    def test_first_run_creates_default_config(self):
        self.assertEqual({**self.saved_config(), "equalizer_apo_config_dir": ""}, ms.DEFAULT_CONFIG)

    def test_defaults_not_shared_between_loads(self):
        ms.CONFIG["profiles"].append({"name": "x"})
        self.assertEqual(ms.DEFAULT_CONFIG["profiles"], [])

    def test_missing_keys_are_backfilled_and_extra_kept(self):
        with open(ms.CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump({"hotkey": "ctrl+shift+f12", "custom": 1}, f)
        cfg, warning = ms.load_config()
        self.assertIsNone(warning)
        self.assertEqual(cfg["hotkey"], "ctrl+shift+f12")
        self.assertEqual(cfg["custom"], 1)
        self.assertEqual(cfg["hotkey_out"], "ctrl+alt+o")
        self.assertEqual(self.saved_config(), cfg)

    def test_broken_config_is_moved_aside_and_recreated(self):
        with open(ms.CONFIG_PATH, "w", encoding="utf-8") as f:
            f.write("{not json")
        cfg, warning = ms.load_config()
        self.assertEqual(cfg, ms.DEFAULT_CONFIG)
        self.assertIn("поврежден", warning)
        with open(ms.CONFIG_PATH + ".broken", encoding="utf-8") as f:
            self.assertEqual(f.read(), "{not json")

    def test_save_is_atomic_and_leaves_no_temp_files(self):
        ms.CONFIG["selected_mic_1"] = "x"
        ms.save_config()
        self.assertEqual(self.saved_config()["selected_mic_1"], "x")
        self.assertEqual(sorted(os.listdir(self.dir)), ["config.json"])

    def test_paths_and_roles_applied(self):
        self.assertEqual(ms.SOUNDVOLUMEVIEW_EXE, os.path.join(self.dir, "SoundVolumeView.exe"))
        self.assertEqual(ms.LOG_FILE, os.path.join(self.dir, "mic_switcher.log"))
        self.assertEqual(ms.ROLES, ad.ALL_ROLES)
        ms.apply_config({**ms.DEFAULT_CONFIG, "set_default_mode": "2"})
        self.assertEqual(ms.ROLES, (ad.ROLE_COMMUNICATIONS,))

    def test_example_config_matches_defaults(self):
        with open(os.path.join(stubs.ROOT, "config.example.json"), encoding="utf-8") as f:
            self.assertEqual(json.load(f), ms.DEFAULT_CONFIG)

    def test_log_rotates_when_too_big(self):
        with open(ms.LOG_FILE, "w", encoding="utf-8") as f:
            f.write("x" * (ms.LOG_MAX_BYTES + 1))
        ms.log("hello")
        self.assertTrue(os.path.isfile(ms.LOG_FILE + ".1"))


class DeviceTests(Base):
    def test_devices_per_kind(self):
        self.assertEqual([d.name for d in ms.get_devices(MIC)], ["Desk Mic", "Headset"])
        self.assertEqual([d.name for d in ms.get_devices(OUT)], ["Динамики", "Наушники"])

    def test_cache_per_flow(self):
        ms.get_devices(MIC)
        ms.display_name(MIC, "{headset}")
        ms.get_devices(OUT)
        ms.display_name(OUT, PHONES.id)
        self.assertEqual(self.backend.list_calls, 2)
        ms.get_devices(MIC, force=True)
        self.tick(ms.DEVICE_CACHE_TTL + 0.1)
        ms.get_devices(OUT)
        self.assertEqual(self.backend.list_calls, 4)

    def test_missing_device_name(self):
        self.assertEqual(ms.display_name(MIC, "{gone}"), "не подключен")

    def test_current_uses_first_role(self):
        self.backend.defaults = {(ad.CAPTURE, 0): "{headset}", (ad.CAPTURE, 2): "{desk}"}
        self.assertEqual(ms.get_current_id(MIC), "{headset}")
        ms.ROLES = (ad.ROLE_COMMUNICATIONS,)
        self.assertEqual(ms.get_current_id(MIC), "{desk}")


class SelectionTests(Base):
    def test_auto_selects_first_two_for_each_kind(self):
        messages = ms.ensure_all_selections()
        self.assertEqual(messages, ["Микрофон: 1) Desk Mic, 2) Headset", "Вывод: 1) Динамики, 2) Наушники"])
        self.assertEqual(self.saved_config()["selected_out_2"], PHONES.id)

    def test_keeps_valid_selection(self):
        self.select_all()
        self.assertEqual(ms.ensure_all_selections(), [None, None])

    def test_duplicate_second_slot_replaced(self):
        ms.CONFIG.update(selected_mic_1="{headset}", selected_mic_2="{headset}")
        ms.ensure_default_selection(MIC)
        self.assertEqual(ms.get_selected(MIC, 2), "{desk}")

    def test_less_than_two_devices(self):
        self.use([HEADSET, PHONES])
        self.assertEqual(
            ms.ensure_all_selections(),
            ["Найдено меньше двух активных микрофонов.", "Найдено меньше двух активных устройств вывода."],
        )

    def test_legacy_ids_migrated_silently(self):
        ms.CONFIG.update(
            selected_mic_1=r"Jabra\Device\Headset\Capture",
            selected_mic_2=r"usb audio\device\desk mic\capture",
            selected_out_1=r"Jabra\Device\Наушники\Render",
            selected_out_2=r"Realtek\Device\Динамики\Render",
        )
        self.assertEqual(ms.ensure_all_selections(), [None, None])
        self.assertEqual(self.saved_config()["selected_mic_1"], "{headset}")
        self.assertEqual(self.saved_config()["selected_out_2"], SPEAKERS.id)

    def test_select_same_as_other_slot_rejected(self):
        ms.CONFIG["selected_out_2"] = SPEAKERS.id
        ms.select_device(OUT, 1, SPEAKERS.id)
        self.assertEqual(ms.get_selected(OUT, 1), "")
        self.assertIn("Out 1, и для Out 2", self.popup_texts()[-1])
        ms.select_device(OUT, 1, PHONES.id)
        self.assertEqual(self.saved_config()["selected_out_1"], PHONES.id)

    def test_selected_label(self):
        ms.CONFIG.update(selected_mic_1="{headset}", selected_mic_2="{gone}")
        self.assertEqual(ms.selected_label(MIC, 1), "Headset")
        self.assertEqual(ms.selected_label(MIC, 2), "не подключен")
        self.assertEqual(ms.selected_label(OUT, 1), "не выбран")


class ToggleTests(Base):
    def setUp(self):
        super().setUp()
        self.select_all()

    def test_pick_toggle_target(self):
        self.assertEqual(ms.pick_toggle_target("a", "a", "b"), (2, "b"))
        self.assertEqual(ms.pick_toggle_target("b", "a", "b"), (1, "a"))
        self.assertEqual(ms.pick_toggle_target(None, "a", "b"), (1, "a"))

    def test_toggle_mic_back_and_forth(self):
        for _ in range(3):
            ms.toggle(MIC)
            self.tick()
        self.assertEqual([c[0] for c in self.backend.set_calls], ["{headset}", "{desk}", "{headset}"])
        self.assertEqual(self.backend.set_calls[0][1], ad.ALL_ROLES)
        self.assertEqual(self.popup_texts()[0], "Переключено на Mic 1: Headset")

    def test_toggle_output(self):
        ms.toggle(OUT)
        self.tick()
        ms.toggle(OUT)
        self.assertEqual([c[0] for c in self.backend.set_calls], [PHONES.id, SPEAKERS.id])
        self.assertEqual(self.popup_texts()[-1], "Переключено на Out 2: Динамики")

    def test_cooldown_is_per_kind(self):
        ms.toggle(MIC)
        self.tick(ms.TOGGLE_COOLDOWN / 2)
        ms.toggle(MIC)
        ms.toggle(OUT)
        self.assertEqual([c[0] for c in self.backend.set_calls], ["{headset}", PHONES.id])

    def test_toggle_requires_both_slots(self):
        ms.CONFIG["selected_out_2"] = ""
        ms.toggle(OUT)
        self.assertEqual(self.backend.set_calls, [])
        self.assertIn("Out 1 и Out 2", self.popup_texts()[-1])

    def test_toggle_skipped_while_switching(self):
        with ms.switch_lock:
            ms.toggle(MIC)
        self.assertEqual(self.backend.set_calls, [])

    def test_error_reported_and_lock_released(self):
        self.backend.fail = True
        ms.toggle(MIC)
        self.assertIn("Не удалось переключить микрофон на: Headset", self.popup_texts()[-1])
        self.assertFalse(ms.switch_lock.locked())

    def test_activate_slot_uses_configured_roles(self):
        ms.ROLES = (ad.ROLE_COMMUNICATIONS,)
        ms.activate_slot(OUT, 2)
        self.assertEqual(self.backend.set_calls, [(SPEAKERS.id, (ad.ROLE_COMMUNICATIONS,))])
        self.assertEqual(self.popup_texts()[-1], "Активен Out 2: Динамики")

    def test_current_summary(self):
        self.backend.defaults[(ad.CAPTURE, 0)] = "{desk}"
        self.backend.defaults[(ad.RENDER, 0)] = "{other}"
        ms.CONFIG.update(active_profile="Стрим", eq_out="Бас")
        self.assertEqual(
            ms.current_summary(),
            "Микрофон: Mic 2, Desk Mic — USB Audio\nВывод: не определен\nПрофиль: Стрим\nEQ: вывод Бас, микрофон Выкл",
        )


class EqualizerTests(Base):
    def setUp(self):
        super().setUp()
        self.select_all()
        self.backend.defaults[(ad.RENDER, 0)] = PHONES.id
        self.backend.defaults[(ad.CAPTURE, 0)] = "{headset}"

    def test_nothing_written_when_off_and_not_installed(self):
        self.assertIsNone(ms.apply_equalizer())
        self.assertFalse(os.path.exists(self.eq_dir))

    def test_warning_when_preset_set_but_not_installed(self):
        ms.CONFIG["eq_out"] = "Бас"
        self.assertIn("Equalizer APO не найден", ms.apply_equalizer())

    def test_set_eq_writes_settings_for_current_devices(self):
        self.install_eq()
        ms.set_eq(OUT, "Бас")
        ms.set_eq(MIC, "Голос")
        text = self.eq_text()
        self.assertIn("Device: 11111111-1111-1111-1111-111111111111", text)
        self.assertIn("Device: Headset", text)
        self.assertIn("Filter: ON LSC Fc 105 Hz Gain 6 dB Q 0.7", text)
        self.assertEqual(self.saved_config()["eq_mic"], "Голос")
        with open(os.path.join(self.eq_dir, "config.txt"), encoding="utf-8") as f:
            self.assertIn("Include: MicSwitcher.txt", f.read())

    def test_eq_follows_output_switch(self):
        self.install_eq()
        ms.set_eq(OUT, "Бас")
        ms.toggle(OUT)
        self.assertIn("Device: 22222222-2222-2222-2222-222222222222", self.eq_text())

    def test_off_clears_previous_settings(self):
        self.install_eq()
        ms.set_eq(OUT, "Бас")
        ms.set_eq(OUT, "off")
        self.assertNotIn("Preamp", self.eq_text())

    def test_custom_preset(self):
        self.install_eq()
        ms.CONFIG["eq_custom_presets"] = {"Мой": {"target": "out", "lines": ["Preamp: -1 dB"]}}
        ms.set_eq(OUT, "Мой")
        self.assertIn("Preamp: -1 dB", self.eq_text())

    def test_noise_suppression_with_plugin(self):
        self.install_eq()
        plugin = os.path.join(os.path.dirname(self.eq_dir), "VSTPlugins", "win-x64", "rnnoise_stereo.dll")
        os.makedirs(os.path.dirname(plugin))
        open(plugin, "w").close()
        ms.set_eq(MIC, "Голос + шумодав")
        text = self.eq_text()
        self.assertIn(f'VSTPlugin: Library "{plugin}"', text)
        self.assertLess(text.index("VSTPlugin"), text.index("Filter: ON HP Fc 80 Hz"))
        self.assertEqual(self.popup_texts()[-1], "Эквалайзер (микрофон): Голос + шумодав")

    def test_noise_threshold(self):
        self.install_eq()
        plugin = os.path.join(os.path.dirname(self.eq_dir), "VSTPlugins", "rnnoise_stereo.dll")
        os.makedirs(os.path.dirname(plugin))
        open(plugin, "w").close()

        ms.set_rnnoise_threshold(90)
        self.assertIn("Применится, когда у микрофона пресет", self.popup_texts()[-1])
        self.assertEqual(self.saved_config()["rnnoise_vad_threshold"], 90)

        ms.set_eq(MIC, "Шумодав")
        self.assertIn(f'ChunkData "{ms.equalizer.rnnoise_chunk(90)}"', self.eq_text())

        ms.set_rnnoise_threshold(None)
        self.assertEqual(self.popup_texts()[-1], "Порог шумодава: По умолчанию плагина (60%)")
        self.assertNotIn("ChunkData", self.eq_text())

    def run_sync(self):
        p = mock.patch.object(ms, "start_background", side_effect=lambda fn, *a: fn(*a))
        p.start()
        self.addCleanup(p.stop)

    def fake_install(self, app_dir, config_dir):
        path = os.path.join(os.path.dirname(config_dir), "VSTPlugins", "RNNoise", "rnnoise_stereo.dll")
        os.makedirs(os.path.dirname(path), exist_ok=True)
        open(path, "w").close()
        return path

    def test_noise_preset_installs_plugin_automatically(self):
        self.install_eq()
        self.run_sync()
        with mock.patch("installer.install_rnnoise", side_effect=self.fake_install) as install:
            ms.set_eq(MIC, "Шумодав")
        install.assert_called_once_with(ms.APP_DIR, self.eq_dir)
        self.assertIn('rnnoise_stereo.dll"', self.eq_text())
        self.assertEqual(self.popup_texts()[-1], "Шумодав установлен.")

    def test_noise_install_failure_keeps_other_presets(self):
        self.install_eq()
        self.run_sync()
        ms.set_eq(OUT, "Бас")
        with mock.patch("installer.install_rnnoise", side_effect=installer_error("Установка отменена")):
            ms.set_eq(MIC, "Шумодав")
        self.assertIn("Не удалось установить шумодав: Установка отменена", self.popup_texts()[-1])
        self.assertNotIn("VSTPlugin", self.eq_text())
        self.assertIn("Бас", self.eq_text())
        self.assertIn("плагин RNNoise не установлен", ms.apply_equalizer())

    def test_noise_preset_without_equalizer_apo(self):
        self.run_sync()
        with mock.patch("installer.install_rnnoise") as install:
            ms.set_eq(MIC, "Шумодав")
        install.assert_not_called()
        self.assertIn("Установить Equalizer APO", self.popup_texts()[-1])

    def test_install_equalizer_apo_runs_installer(self):
        with mock.patch("installer.download_equalizer_apo", return_value=r"C:\tmp\EqualizerAPO.exe"), \
                mock.patch.object(ms.os, "startfile", create=True) as start:
            ms.install_equalizer_apo()
        start.assert_called_once_with(r"C:\tmp\EqualizerAPO.exe")
        self.assertIn("Configurator", self.popup_texts()[-1])

    def test_parallel_install_refused(self):
        with ms.install_lock:
            ms.install_noise_suppression()
        self.assertEqual(self.popup_texts()[-1], "Установка уже идет.")

    def test_write_error_reported(self):
        self.install_eq()
        with mock.patch("equalizer.install", side_effect=PermissionError("нет доступа")):
            ms.set_eq(OUT, "Бас")
        self.assertIn("Не удалось записать настройки Equalizer APO", self.popup_texts()[-1])


class ProfileTests(Base):
    def setUp(self):
        super().setUp()
        self.select_all()
        self.install_eq()
        ms.CONFIG["profiles"] = [
            {"name": "Гарнитура", "mic": "{headset}", "out": PHONES.id, "eq_out": "Бас", "eq_mic": "Голос", "hotkey": "ctrl+alt+1"},
            {"name": "Колонки", "mic": "{desk}", "out": SPEAKERS.id},
        ]

    def test_apply_profile(self):
        ms.apply_profile("гарнитура")
        self.assertEqual([c[0] for c in self.backend.set_calls], [PHONES.id, "{headset}"])
        self.assertEqual(ms.CONFIG["eq_out"], "Бас")
        self.assertEqual(self.saved_config()["active_profile"], "Гарнитура")
        self.assertEqual(self.popup_texts()[-1], "Профиль «Гарнитура»\nВывод: Наушники\nМикрофон: Headset")
        self.assertIn("Device: 11111111-1111-1111-1111-111111111111", self.eq_text())

    def test_profile_without_eq_keeps_presets(self):
        ms.CONFIG["eq_out"] = "Мягкий"
        ms.apply_profile("Колонки")
        self.assertEqual(ms.CONFIG["eq_out"], "Мягкий")

    def test_missing_device_skipped(self):
        self.backend = self.use([DESK, PHONES, SPEAKERS])
        ms.apply_profile("Гарнитура")
        self.assertEqual([c[0] for c in self.backend.set_calls], [PHONES.id])
        self.assertIn("Микрофон не подключен", self.popup_texts()[-1])

    def test_next_profile_cycles(self):
        ms.next_profile()
        ms.next_profile()
        ms.next_profile()
        self.assertEqual(ms.CONFIG["active_profile"], "Гарнитура")
        self.assertEqual(len(self.backend.set_calls), 6)

    def test_next_profile_without_profiles(self):
        ms.CONFIG["profiles"] = []
        ms.next_profile()
        self.assertIn("Профилей нет", self.popup_texts()[-1])

    def test_save_current_asks_name_with_default(self):
        self.backend.defaults[(ad.RENDER, 0)] = SPEAKERS.id
        self.backend.defaults[(ad.CAPTURE, 0)] = "{desk}"
        ms.save_current_profile()
        title, prompt, initial, callback = ms.popups.ask.call_args.args
        self.assertEqual(initial, "Динамики + Desk Mic")
        self.assertIs(callback, ms.save_profile_named)

    def test_save_new_and_overwrite_keeps_hotkey(self):
        self.backend.defaults[(ad.RENDER, 0)] = SPEAKERS.id
        self.backend.defaults[(ad.CAPTURE, 0)] = "{desk}"
        ms.CONFIG["eq_out"] = "Игры: шаги"
        ms.save_profile_named(" Стрим ")
        ms.save_profile_named("гарнитура")
        saved = {p["name"]: p for p in self.saved_config()["profiles"]}
        self.assertEqual(saved["Стрим"]["out"], SPEAKERS.id)
        self.assertEqual(saved["Стрим"]["eq_out"], "Игры: шаги")
        self.assertEqual(saved["гарнитура"]["mic"], "{desk}")
        self.assertEqual(saved["гарнитура"]["hotkey"], "ctrl+alt+1")
        self.assertEqual(ms.CONFIG["active_profile"], "гарнитура")
        self.assertIn("обновлен", self.popup_texts()[-1])

    def test_cancel_name_does_nothing(self):
        ms.save_profile_named(None)
        ms.save_profile_named("  ")
        self.assertEqual(len(ms.CONFIG["profiles"]), 2)

    def test_delete_profile(self):
        ms.CONFIG["active_profile"] = "Колонки"
        ms.delete_profile("Колонки")
        self.assertEqual([p["name"] for p in self.saved_config()["profiles"]], ["Гарнитура"])
        self.assertEqual(ms.CONFIG["active_profile"], "")


class HotkeyTests(Base):
    def test_bindings(self):
        ms.CONFIG["profiles"] = [{"name": "Гарнитура", "hotkey": "ctrl+alt+1"}, {"name": "Без"}]
        ms.CONFIG["hotkey_out"] = ""
        combos = [b[0] for b in ms.hotkey_bindings()]
        self.assertEqual(combos, ["ctrl+alt+m", "ctrl+alt+p", "ctrl+alt+1"])

    def test_register_replaces_old_and_collects_errors(self):
        kb = ms.keyboard
        kb.add_hotkey.side_effect = lambda combo, fn, args: combo
        self.assertEqual(ms.register_hotkeys(), [])
        self.assertEqual(ms.hotkey_handles, ["ctrl+alt+m", "ctrl+alt+o", "ctrl+alt+p"])

        def add(combo, fn, args):
            if combo == "ctrl+alt+o":
                raise ValueError("bad")
            return combo

        kb.add_hotkey.side_effect = add
        errors = ms.register_hotkeys()
        self.assertEqual(kb.remove_hotkey.call_count, 3)
        self.assertEqual(len(errors), 1)
        self.assertIn("ctrl+alt+o (переключить вывод)", errors[0])


class WatchTests(Base):
    def test_diff_devices(self):
        before = {d.id: d for d in (HEADSET, DESK)}
        after = {d.id: d for d in (DESK, CAMERA)}
        self.assertEqual(ms.diff_devices(MIC, None, after), [])
        self.assertEqual(ms.diff_devices(MIC, after, dict(after)), [])
        self.assertEqual(
            ms.diff_devices(MIC, before, after),
            ["Подключен микрофон: Camera — Logi", "Отключен микрофон: Headset — Jabra"],
        )

    def run_watch(self, states):
        stop = mock.Mock()

        def wait(timeout):
            if not states:
                return True
            states.pop(0)()
            return False

        stop.wait.side_effect = wait
        ms.watch_devices(stop)

    def test_reports_changes_and_keeps_selection(self):
        self.select_all()

        def unplug():
            self.backend.devices = [DESK, SPEAKERS]

        self.run_watch([lambda: None, unplug])
        self.assertEqual(self.popup_texts(), ["Отключен микрофон: Headset — Jabra\nОтключен вывод: Наушники — Jabra"])
        self.assertEqual(ms.get_selected(MIC, 1), "{headset}")

    def test_eq_follows_default_changed_outside(self):
        self.install_eq()
        ms.CONFIG["eq_out"] = "Бас"
        self.backend.defaults[(ad.RENDER, 0)] = PHONES.id

        def windows_switch():
            self.backend.defaults[(ad.RENDER, 0)] = SPEAKERS.id

        self.run_watch([lambda: None, windows_switch])
        self.assertIn("Device: 22222222-2222-2222-2222-222222222222", self.eq_text())
        self.assertEqual(self.popup_texts(), [])

    def test_refresh_single_popup(self):
        self.use([HEADSET, DESK, CAMERA, PHONES, SPEAKERS])
        ms.refresh_devices()
        texts = self.popup_texts()
        self.assertEqual(len(texts), 1)
        self.assertIn("Микрофон: найдено 3", texts[0])
        self.assertIn("Вывод: 1) Динамики, 2) Наушники", texts[0])

    def test_watcher_not_started_for_slow_backend(self):
        self.backend.watchable = False
        with mock.patch("threading.Thread") as thread:
            ms.start_watcher()
        thread.assert_not_called()


class PanelStateTests(Base):
    def setUp(self):
        super().setUp()
        self.select_all()
        self.backend.defaults[(ad.CAPTURE, 0)] = "{desk}"
        self.backend.defaults[(ad.RENDER, 0)] = PHONES.id
        p = mock.patch.object(ms, "is_autorun_enabled", return_value=True)
        p.start()
        self.addCleanup(p.stop)

    def test_state(self):
        ms.CONFIG["profiles"] = [{"name": "Стрим", "hotkey": "ctrl+alt+1"}]
        ms.CONFIG.update(active_profile="Стрим", eq_mic="Голос")
        state = ms.panel_state()
        mic, out = state["kinds"]
        self.assertEqual(mic["key"], "mic")
        self.assertEqual(mic["current"], "Mic 2, Desk Mic — USB Audio")
        self.assertEqual(mic["slots"], {1: "{headset}", 2: "{desk}"})
        self.assertEqual([d["label"] for d in out["devices"]], ["Динамики — Realtek", "Наушники — Jabra"])
        self.assertEqual(out["hotkey"], "ctrl+alt+o")
        self.assertEqual(state["profiles"], [{"name": "Стрим", "hotkey": "ctrl+alt+1"}])
        self.assertEqual(state["profile_name_suggestion"], "Наушники + Desk Mic")
        self.assertFalse(state["eq"]["found"])
        self.assertEqual(state["eq"]["kinds"]["mic"]["current"], "Голос")
        self.assertEqual(state["eq"]["kinds"]["out"]["presets"][0], ("off", "Выкл"))
        self.assertEqual(state["eq"]["threshold"], None)
        self.assertEqual(state["hotkeys"], {"hotkey": "ctrl+alt+m", "hotkey_out": "ctrl+alt+o", "hotkey_profile": "ctrl+alt+p"})
        self.assertTrue(state["autorun"])
        json.dumps(state)

    def test_set_hotkeys(self):
        ms.keyboard.add_hotkey.side_effect = lambda combo, fn, args: combo
        ms.set_hotkeys({"hotkey": " ctrl+shift+m ", "hotkey_out": "", "hotkey_profile": "ctrl+alt+p"})
        self.assertEqual(self.saved_config()["hotkey"], "ctrl+shift+m")
        self.assertEqual(ms.hotkey_handles, ["ctrl+shift+m", "ctrl+alt+p"])
        self.assertEqual(self.popup_texts()[-1], "Хоткеи сохранены.")

    def test_set_profile_hotkey(self):
        ms.CONFIG["profiles"] = [{"name": "Стрим", "mic": "{desk}"}]
        ms.set_profile_hotkey("стрим", "ctrl+alt+9")
        saved = self.saved_config()["profiles"][0]
        self.assertEqual((saved["name"], saved["mic"], saved["hotkey"]), ("Стрим", "{desk}", "ctrl+alt+9"))
        ms.set_profile_hotkey("нет такого", "x")

    def test_set_auto_refresh(self):
        with mock.patch("threading.Thread") as thread:
            ms.set_auto_refresh(False)
            thread.assert_not_called()
            ms.set_auto_refresh(True)
            thread.assert_called_once()
        self.assertTrue(self.saved_config()["auto_refresh"])

    def test_panel_actions_map_kind_keys(self):
        actions = ms.panel_actions()
        actions.toggle("out")
        self.assertEqual(self.backend.set_calls[0][0], SPEAKERS.id)
        actions.set_eq("mic", "Радио")
        self.assertEqual(ms.CONFIG["eq_mic"], "Радио")


class AutorunTests(Base):
    def test_frozen_command(self):
        with mock.patch.object(sys, "frozen", True, create=True), \
                mock.patch.object(sys, "executable", r"C:\App\MicSwitcher.exe"):
            self.assertEqual(ms.get_autorun_command(), r'"C:\App\MicSwitcher.exe"')

    def test_source_command_prefers_pythonw(self):
        exe = os.path.join(self.dir, "python.exe")
        pythonw = os.path.join(self.dir, "pythonw.exe")
        open(pythonw, "w").close()
        with mock.patch.object(sys, "executable", exe):
            self.assertTrue(ms.get_autorun_command().startswith(f'"{pythonw}" "'))


if __name__ == "__main__":
    unittest.main()
