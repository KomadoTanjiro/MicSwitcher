"""Панель на настоящем tkinter: построение, отрисовка снимка, действия.
Фоновые потоки панели выполняются сразу, чтобы тест был детерминированным."""
import tkinter as tk
import unittest
from unittest import mock

from tests import stubs  # noqa: F401

import panel as panel_ui


def make_state(**overrides):
    state = {
        "backend": "Core Audio",
        "watchable": True,
        "kinds": [
            {
                "key": "mic",
                "title": "Микрофон",
                "short": "Mic",
                "devices": [{"id": "{a}", "label": "Headset — Jabra"}, {"id": "{b}", "label": "Desk Mic — USB"}],
                "slots": {1: "{a}", 2: "{gone}"},
                "current": "Mic 1, Headset — Jabra",
                "hotkey": "ctrl+alt+m",
            },
            {
                "key": "out",
                "title": "Вывод",
                "short": "Out",
                "devices": [{"id": "{o1}", "label": "Наушники"}, {"id": "{o2}", "label": "Колонки"}],
                "slots": {1: "{o1}", 2: "{o2}"},
                "current": "не определен",
                "hotkey": "",
            },
        ],
        "profiles": [{"name": "Игры", "hotkey": "ctrl+alt+1"}, {"name": "Созвон", "hotkey": ""}],
        "active_profile": "Игры",
        "profile_name_suggestion": "Наушники + Headset",
        "eq": {
            "found": True,
            "dir": r"C:\Program Files\EqualizerAPO\config",
            "download_url": "https://eq",
            "rnnoise_url": "https://rn",
            "rnnoise": None,
            "kinds": {
                "out": {"presets": [("off", "Выкл"), ("Бас", "Бас")], "current": "Бас"},
                "mic": {"presets": [("off", "Выкл"), ("Шумодав", "Шумодав")], "current": "off"},
            },
            "thresholds": [(None, "По умолчанию плагина (60%)"), (90, "90% (рекомендуется)")],
            "threshold": None,
        },
        "hotkeys": {"hotkey": "ctrl+alt+m", "hotkey_out": "ctrl+alt+o", "hotkey_profile": ""},
        "autorun": True,
        "auto_refresh": False,
    }
    state.update(overrides)
    return state


class SyncThread:
    def __init__(self, target, args=(), daemon=None):
        self.target, self.args = target, args

    def start(self):
        self.target(*self.args)


class PanelTests(unittest.TestCase):
    def setUp(self):
        try:
            self.root = tk.Tk()
        except tk.TclError as e:
            self.skipTest(f"нет дисплея: {e}")
        self.root.withdraw()
        self.addCleanup(self.root.destroy)

        p = mock.patch.object(panel_ui.threading, "Thread", SyncThread)
        p.start()
        self.addCleanup(p.stop)

        self.current = make_state()
        self.actions = mock.Mock()
        self.actions.state.side_effect = lambda: self.current
        self.actions.call_ui.side_effect = lambda fn, *args: fn(*args)
        self.actions.record_hotkey.return_value = "ctrl+shift+k"
        self.panel = panel_ui.Panel(self.root, self.actions)
        self.root.update_idletasks()

    def choose(self, choice, index):
        choice.box.current(index)
        choice.box.event_generate("<<ComboboxSelected>>")

    def test_renders_state(self):
        mic = self.panel.kind_widgets["mic"]
        self.assertEqual(mic["current"].cget("text"), "Сейчас: Mic 1, Headset — Jabra")
        self.assertEqual(mic["slots"][1].box.get(), "Headset — Jabra")
        self.assertEqual(mic["slots"][2].box.get(), "не подключен")
        self.assertEqual(mic["toggle"].cget("text"), "Переключить  (ctrl+alt+m)")
        self.assertEqual(self.panel.kind_widgets["out"]["toggle"].cget("text"), "Переключить")
        self.assertEqual(self.panel.profile_list.get(0, "end"), ("• Игры   (ctrl+alt+1)", "   Созвон"))
        self.assertEqual(self.panel.profile_name.get(), "Наушники + Headset")
        self.assertEqual(self.panel.eq_choices["out"].box.get(), "Бас")
        self.assertEqual(self.panel.threshold.box.get(), "По умолчанию плагина (60%)")
        self.assertEqual(self.panel.hotkey_entries["hotkey_out"].get(), "ctrl+alt+o")
        self.assertTrue(self.panel.autorun_var.get())
        self.assertFalse(self.panel.watch_var.get())
        self.assertIn("Поставится сам", self.panel.rnnoise_status.cget("text"))

    def test_device_and_eq_choices_call_actions(self):
        self.choose(self.panel.kind_widgets["mic"]["slots"][2], 1)
        self.actions.select_device.assert_called_with("mic", 2, "{b}")
        self.choose(self.panel.eq_choices["mic"], 1)
        self.actions.set_eq.assert_called_with("mic", "Шумодав")
        self.choose(self.panel.threshold, 1)
        self.actions.set_rnnoise_threshold.assert_called_with(90)
        self.choose(self.panel.threshold, 0)
        self.actions.set_rnnoise_threshold.assert_called_with(None)

    def test_buttons(self):
        self.panel.kind_widgets["out"]["toggle"].invoke()
        self.actions.toggle.assert_called_with("out")
        self.panel.profile_list.selection_set(1)
        self.panel._profile_apply()
        self.actions.apply_profile.assert_called_with("Созвон")
        self.panel.eq_button.invoke()
        self.actions.open_target.assert_called_with(r"C:\Program Files\EqualizerAPO\config")

    def test_panel_stays_open_and_refreshes_after_action(self):
        before = self.actions.state.call_count
        self.current = make_state(active_profile="Созвон")
        self.panel.kind_widgets["mic"]["toggle"].invoke()
        self.assertGreater(self.actions.state.call_count, before)
        self.assertEqual(self.panel.profile_list.get(1), "• Созвон")
        self.assertFalse(self.panel.closed)

    def test_save_profile_and_hotkeys(self):
        self.panel.profile_name.delete(0, "end")
        self.panel.profile_name.insert(0, " Стрим ")
        self.panel._profile_save()
        self.actions.save_profile_named.assert_called_with("Стрим")

        entry = self.panel.hotkey_entries["hotkey_profile"]
        entry.insert(0, "ctrl+alt+p")
        self.panel._save_hotkeys()
        self.actions.set_hotkeys.assert_called_with(
            {"hotkey": "ctrl+alt+m", "hotkey_out": "ctrl+alt+o", "hotkey_profile": "ctrl+alt+p"}
        )

    def test_typed_hotkey_survives_periodic_refresh(self):
        entry = self.panel.hotkey_entries["hotkey"]
        entry.delete(0, "end")
        entry.insert(0, "ctrl+shift+x")
        self.panel.request_refresh()
        self.assertEqual(entry.get(), "ctrl+shift+x")

    def test_record_hotkey(self):
        entry = self.panel.hotkey_entries["hotkey_out"]
        button = mock.Mock()
        self.panel._record(entry, button)
        self.assertEqual(entry.get(), "ctrl+shift+k")
        button.configure.assert_called_with(text="Записать", state="normal")

    def test_profile_hotkey(self):
        self.panel.profile_list.selection_set(0)
        self.panel._profile_selected()
        self.assertEqual(self.panel.profile_hotkey.get(), "ctrl+alt+1")
        self.panel.profile_hotkey.delete(0, "end")
        self.panel._profile_set_hotkey()
        self.actions.set_profile_hotkey.assert_called_with("Игры", "")

    def test_close(self):
        self.panel.close()
        self.assertTrue(self.panel.closed)
        self.panel.request_refresh()


if __name__ == "__main__":
    unittest.main()
