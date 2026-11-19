import base64
import os
import struct
import tempfile
import unittest
import xml.etree.ElementTree as ET

from tests import stubs  # noqa: F401

import equalizer as eq


class PresetTests(unittest.TestCase):
    def test_builtin_presets_per_target(self):
        self.assertIn("Бас", eq.presets("out"))
        self.assertIn("Голос", eq.presets("mic"))
        self.assertNotIn("Голос", eq.presets("out"))

    def test_builtin_commands_look_like_equalizer_apo(self):
        for target in ("out", "mic"):
            for name, lines in eq.presets(target).items():
                for line in lines:
                    self.assertRegex(
                        line,
                        r'^(Preamp: -?\d+ dB|Filter: ON [A-Z]+ Fc \d+ Hz( Gain -?\d+ dB)?( Q [\d.]+)?|VSTPlugin: Library "\{rnnoise\}")$',
                        name,
                    )

    def test_custom_presets(self):
        custom = {
            "Мой": {"target": "out", "lines": ["Preamp: -2 dB", " "]},
            "Пустой": {"target": "out", "lines": []},
            "Чужой": {"target": "mic", "lines": ["Preamp: -1 dB"]},
            "Битый": "строка",
        }
        out = eq.presets("out", custom)
        self.assertEqual(out["Мой"], ["Preamp: -2 dB"])
        self.assertNotIn("Пустой", out)
        self.assertNotIn("Чужой", out)
        self.assertNotIn("Битый", out)
        self.assertIn("Бас", out)

    def test_preset_label(self):
        self.assertEqual(eq.preset_label("off"), "Выкл")
        self.assertEqual(eq.preset_label(None), "Выкл")
        self.assertEqual(eq.preset_label("Бас"), "Бас")


class RenderTests(unittest.TestCase):
    def test_device_pattern_uses_endpoint_guid(self):
        guid = "{e6ea5b0c-1234-4abc-9def-0123456789ab}"
        self.assertEqual(eq.device_pattern("{0.0.0.00000000}." + guid, "Динамики"), guid.strip("{}"))
        self.assertEqual(eq.device_pattern(r"Realtek\Device\Динамики\Render", "Динамики"), "Динамики")

    def test_render(self):
        text = eq.render([("Вывод", "Наушники — Jabra", "{g}", "Бас", ["Preamp: -6 dB"])])
        lines = text.splitlines()
        self.assertTrue(lines[0].startswith("# MicSwitcher"))
        self.assertIn("Device: {g}", lines)
        self.assertLess(lines.index("Device: {g}"), lines.index("Preamp: -6 dB"))
        self.assertEqual(lines[-1], "Device: all")

    def test_render_empty_resets(self):
        self.assertIn("Device: all", eq.render([]))
        self.assertNotIn("Preamp", eq.render([]))


class RnnoiseTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.config_dir = os.path.join(self.tmp.name, "EqualizerAPO", "config")
        self.plugins = os.path.join(self.tmp.name, "EqualizerAPO", "VSTPlugins")
        os.makedirs(self.config_dir)

    def put(self, *parts):
        path = os.path.join(self.plugins, *parts)
        os.makedirs(os.path.dirname(path), exist_ok=True)
        open(path, "w").close()
        return path

    def test_noise_presets_need_plugin(self):
        mic = eq.presets("mic")
        self.assertTrue(eq.needs_rnnoise(mic["Шумодав"]))
        self.assertTrue(eq.needs_rnnoise(mic["Голос + шумодав"]))
        self.assertFalse(eq.needs_rnnoise(mic["Голос"]))

    def test_resolve_commands(self):
        commands = eq.presets("mic")["Голос + шумодав"]
        self.assertIsNone(eq.resolve_commands(commands, None))
        resolved = eq.resolve_commands(commands, r"C:\Program Files\EqualizerAPO\VSTPlugins\rnnoise_stereo.dll")
        self.assertEqual(resolved[0], r'VSTPlugin: Library "C:\Program Files\EqualizerAPO\VSTPlugins\rnnoise_stereo.dll"')
        self.assertEqual(eq.resolve_commands(["Preamp: -1 dB"], None), ["Preamp: -1 dB"])

    def test_threshold_parsing_and_labels(self):
        self.assertIsNone(eq.parse_threshold(None))
        self.assertIsNone(eq.parse_threshold(""))
        self.assertIsNone(eq.parse_threshold(True))
        self.assertIsNone(eq.parse_threshold(150))
        self.assertIsNone(eq.parse_threshold("abc"))
        self.assertEqual(eq.parse_threshold("90"), 90)
        self.assertEqual(eq.threshold_label(None), "По умолчанию плагина (60%)")
        self.assertEqual(eq.threshold_label(90), "90% (рекомендуется)")
        self.assertEqual(eq.threshold_label(0), "0%: только шумоподавление")

    def test_chunk_is_juce_binary_xml(self):
        blob = base64.b64decode(eq.rnnoise_chunk(85))
        magic, length = struct.unpack("<II", blob[:8])
        self.assertEqual(magic, 0x21324356)
        self.assertEqual(length, len(blob) - 9)
        self.assertEqual(blob[-1:], b"\0")
        root = ET.fromstring(blob[8:8 + length].decode("utf-8"))
        self.assertEqual(root.tag, "RNNoise")
        params = {p.get("id"): p.get("value") for p in root.findall("PARAM")}
        self.assertEqual(
            params,
            {"vad_threshold": "0.85", "vad_grace_period": "20", "vad_retroactive_grace_period": "0"},
        )

    def test_resolve_with_threshold(self):
        commands = eq.presets("mic")["Шумодав"]
        line = eq.resolve_commands(commands, r"C:\p\rnnoise_stereo.dll", 90)[0]
        self.assertEqual(line, f'VSTPlugin: Library "C:\\p\\rnnoise_stereo.dll" ChunkData "{eq.rnnoise_chunk(90)}"')
        self.assertNotIn("ChunkData", eq.resolve_commands(commands, "x.dll", None)[0])

    def test_find_prefers_stereo_x64(self):
        self.put("win-x86", "vst", "rnnoise_stereo.dll")
        self.put("win-x64", "vst", "rnnoise_mono.dll")
        best = self.put("win-x64", "vst", "rnnoise_stereo.dll")
        self.assertEqual(eq.find_rnnoise(self.config_dir), best)

    def test_find_case_insensitive_and_missing(self):
        self.assertIsNone(eq.find_rnnoise(self.config_dir))
        path = self.put("RNNoise_Stereo.dll")
        self.assertEqual(eq.find_rnnoise(self.config_dir), path)
        self.assertIsNone(eq.find_rnnoise(None))

    def test_override(self):
        path = self.put("custom.dll")
        self.assertEqual(eq.find_rnnoise(None, override=path), path)
        self.assertIsNone(eq.find_rnnoise(self.config_dir, override=path + ".nope"))


class InstallTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.dir = self.tmp.name
        self.config_txt = os.path.join(self.dir, "config.txt")

    def read(self, name):
        with open(os.path.join(self.dir, name), encoding="utf-8") as f:
            return f.read()

    def test_include_appended_once_and_user_config_kept(self):
        with open(self.config_txt, "w", encoding="utf-8") as f:
            f.write("Preamp: -1 dB\nInclude: example.txt")
        self.assertTrue(eq.ensure_include(self.dir))
        self.assertFalse(eq.ensure_include(self.dir))
        self.assertEqual(self.read("config.txt"), "Preamp: -1 dB\nInclude: example.txt\nInclude: MicSwitcher.txt\n")

    def test_include_detected_case_insensitive(self):
        with open(self.config_txt, "w", encoding="utf-8") as f:
            f.write("include:   micswitcher.txt\n")
        self.assertFalse(eq.ensure_include(self.dir))

    def test_include_creates_config(self):
        eq.ensure_include(self.dir)
        self.assertEqual(self.read("config.txt"), "Include: MicSwitcher.txt\n")

    def test_install_writes_settings_without_temp_files(self):
        self.assertFalse(eq.settings_exist(self.dir))
        eq.install(self.dir, "Device: all\n")
        self.assertTrue(eq.settings_exist(self.dir))
        self.assertEqual(self.read("MicSwitcher.txt"), "Device: all\n")
        self.assertEqual(sorted(os.listdir(self.dir)), ["MicSwitcher.txt", "config.txt"])


class FindConfigDirTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)

    def test_override(self):
        self.assertEqual(eq.find_config_dir(self.tmp.name, environ={}), self.tmp.name)
        self.assertIsNone(eq.find_config_dir(os.path.join(self.tmp.name, "nope"), environ={}))

    def test_program_files(self):
        path = os.path.join(self.tmp.name, "EqualizerAPO", "config")
        os.makedirs(path)
        self.assertEqual(eq.find_config_dir(environ={"ProgramFiles": self.tmp.name}), path)

    def test_registry_first(self):
        tmp = self.tmp.name

        class Key:
            def __enter__(self):
                return self

            def __exit__(self, *a):
                return False

        class Reg:
            HKEY_LOCAL_MACHINE = "HKLM"

            @staticmethod
            def OpenKey(root, path):
                assert (root, path) == ("HKLM", eq.REG_PATH)
                return Key()

            @staticmethod
            def QueryValueEx(key, name):
                assert name == "ConfigPath"
                return tmp, 1

        self.assertEqual(eq.find_config_dir(winreg=Reg, environ={}), tmp)

    def test_not_installed(self):
        self.assertIsNone(eq.find_config_dir(environ={"ProgramFiles": self.tmp.name}))


if __name__ == "__main__":
    unittest.main()
