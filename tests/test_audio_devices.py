import os
import sys
import tempfile
import types
import unittest
from unittest import mock

from tests import stubs  # noqa: F401

import audio_devices as ad
from audio_devices.svv import COLUMNS, SvvBackend

HEADER = ",".join(COLUMNS)


def row(name, dev_id, default="", multimedia="", comms="", direction="Capture", kind="Device", state="Active"):
    return ",".join([name, direction, kind, state, default, comms, multimedia, dev_id])


HEADSET = row("Headset", r"Jabra\Device\Microphone\Capture", default="Capture", multimedia="Capture")
DESK = row("Desk Mic", r"USB Audio\Device\Microphone\Capture", comms="Capture")
CAMERA = row("Camera", r"Logi\Device\Microphone\Capture")
DISABLED = row("Old", r"Old\Device\Microphone\Capture", state="Disabled")
SPEAKERS = row("Speakers", r"Realtek\Device\Speakers\Render", default="Render", direction="Render")


class FakeRun:
    def __init__(self, rows, returncode=0):
        self.rows = rows
        self.returncode = returncode
        self.calls = []

    def __call__(self, args):
        self.calls.append(args)
        if args[0] == "/scomma":
            with open(args[1], "w", encoding="utf-8-sig") as f:
                f.write("\n".join([HEADER] + self.rows) + "\n")
        return types.SimpleNamespace(returncode=self.returncode, stdout="", stderr="")


class DeviceTests(unittest.TestCase):
    def test_label_with_interface(self):
        self.assertEqual(ad.Device("id", "Микрофон", "Realtek(R) Audio").label, "Микрофон — Realtek(R) Audio")

    def test_label_without_interface(self):
        self.assertEqual(ad.Device("id", "Jabra", "Jabra").label, "Jabra")
        self.assertEqual(ad.Device("id", "", "").label, "id")

    def test_legacy_id_matches_soundvolumeview_format(self):
        self.assertEqual(
            ad.Device("x", "Микрофон", "Realtek(R) Audio").legacy_id,
            r"Realtek(R) Audio\Device\Микрофон\Capture",
        )
        self.assertEqual(
            ad.Device("x", "Динамики", "Realtek(R) Audio", ad.RENDER).legacy_id,
            r"Realtek(R) Audio\Device\Динамики\Render",
        )

    def test_parse_roles(self):
        self.assertEqual(ad.parse_roles("all"), ad.ALL_ROLES)
        self.assertEqual(ad.parse_roles("2"), (ad.ROLE_COMMUNICATIONS,))
        self.assertEqual(ad.parse_roles(" Multimedia "), (ad.ROLE_MULTIMEDIA,))
        self.assertEqual(ad.parse_roles(0), (ad.ROLE_CONSOLE,))
        self.assertEqual(ad.parse_roles("ерунда"), ad.ALL_ROLES)


class OpenBackendTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.svv = os.path.join(self.tmp.name, "SoundVolumeView.exe")

    def make_svv(self):
        open(self.svv, "w").close()

    def without_core(self):
        return mock.patch.dict(sys.modules, {"audio_devices.core": None})

    def with_core(self):
        core = types.ModuleType("audio_devices.core")
        core.CoreAudioBackend = type("CoreAudioBackend", (), {"name": "Core Audio"})
        return mock.patch.dict(sys.modules, {"audio_devices.core": core})

    def test_auto_prefers_core(self):
        self.make_svv()
        with self.with_core():
            backend, warning = ad.open_backend("auto", self.svv)
        self.assertEqual(backend.name, "Core Audio")
        self.assertIsNone(warning)

    def test_auto_falls_back_to_svv(self):
        self.make_svv()
        with self.without_core():
            backend, warning = ad.open_backend("auto", self.svv)
        self.assertIsInstance(backend, SvvBackend)
        self.assertIn("SoundVolumeView", warning)

    def test_auto_without_core_and_svv_fails(self):
        with self.without_core(), self.assertRaises(ad.AudioBackendError):
            ad.open_backend("auto", self.svv)

    def test_core_mode_never_falls_back(self):
        self.make_svv()
        with self.without_core(), self.assertRaises(ad.AudioBackendError):
            ad.open_backend("core", self.svv)

    def test_svv_mode(self):
        self.make_svv()
        with self.with_core():
            backend, _ = ad.open_backend("SVV", self.svv)
        self.assertIsInstance(backend, SvvBackend)

    def test_svv_mode_without_exe(self):
        with self.assertRaises(FileNotFoundError):
            ad.open_backend("svv", self.svv)


class SvvBackendTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        exe = os.path.join(self.tmp.name, "SoundVolumeView.exe")
        open(exe, "w").close()
        self.backend = SvvBackend(exe)
        p = mock.patch("time.sleep")
        p.start()
        self.addCleanup(p.stop)

    def use(self, rows, returncode=0):
        self.backend.run = FakeRun(rows, returncode)
        return self.backend.run

    def test_lists_only_active_devices_of_flow(self):
        self.use([HEADSET, SPEAKERS, DISABLED, DESK, CAMERA])
        devices = self.backend.list_devices(ad.CAPTURE)
        self.assertEqual([d.name for d in devices], ["Camera", "Desk Mic", "Headset"])
        self.assertEqual(devices[2].interface, "Jabra")
        self.assertEqual([d.name for d in self.backend.list_devices(ad.RENDER)], ["Speakers"])

    def test_default_per_role(self):
        self.use([DESK, HEADSET, SPEAKERS])
        self.assertEqual(self.backend.get_default_id(ad.CAPTURE, ad.ROLE_CONSOLE), r"Jabra\Device\Microphone\Capture")
        self.assertEqual(
            self.backend.get_default_id(ad.CAPTURE, ad.ROLE_COMMUNICATIONS), r"USB Audio\Device\Microphone\Capture"
        )
        self.assertEqual(self.backend.get_default_id(ad.RENDER, ad.ROLE_CONSOLE), r"Realtek\Device\Speakers\Render")

    def test_no_default(self):
        self.use([CAMERA])
        self.assertIsNone(self.backend.get_default_id(ad.CAPTURE, ad.ROLE_CONSOLE))

    def test_set_default_all_roles_in_one_call(self):
        run = self.use([])
        self.backend.set_default("id", ad.ALL_ROLES)
        self.assertEqual(run.calls, [["/SetDefault", "id", "all"]])

    def test_set_default_single_role(self):
        run = self.use([])
        self.backend.set_default("id", (ad.ROLE_COMMUNICATIONS,))
        self.assertEqual(run.calls, [["/SetDefault", "id", "2"]])

    def test_errors_raise(self):
        self.use([], returncode=1)
        with self.assertRaises(RuntimeError):
            self.backend.list_devices(ad.CAPTURE)
        with self.assertRaises(RuntimeError):
            self.backend.set_default("id", ad.ALL_ROLES)

    def test_empty_export_and_temp_file_removed(self):
        run = self.use([])
        run.rows = []
        self.assertEqual(self.backend.list_devices(ad.CAPTURE), [])
        self.assertFalse(os.path.exists(run.calls[0][1]))


if __name__ == "__main__":
    unittest.main()
