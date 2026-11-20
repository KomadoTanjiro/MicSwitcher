import base64
import hashlib
import io
import os
import tempfile
import unittest
import zipfile
from unittest import mock

from tests import stubs  # noqa: F401

import installer as inst


class FakeResponse(io.BytesIO):
    def __enter__(self):
        return self

    def __exit__(self, *a):
        self.close()
        return False


def opener_for(data, calls=None):
    def opener(request, timeout=None):
        if calls is not None:
            calls.append(request.full_url)
        return FakeResponse(data)

    return opener


def rnnoise_zip():
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("win-rnnoise/vst/rnnoise_stereo.dll", b"VST-STEREO")
        z.writestr("win-rnnoise/vst/rnnoise_mono.dll", b"VST-MONO")
        z.writestr("win-rnnoise/rnnoise_stereo.lv2/rnnoise_stereo.dll", b"LV2")
    return buf.getvalue()


class Base(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.dir = self.tmp.name
        self.config_dir = os.path.join(self.dir, "EqualizerAPO", "config")
        os.makedirs(self.config_dir)
        self.app_dir = os.path.join(self.dir, "app")
        os.makedirs(self.app_dir)

    def read(self, path):
        with open(path, "rb") as f:
            return f.read()


class DownloadTests(Base):
    def test_hash_ok(self):
        data = b"payload"
        path = inst.download("https://x", hashlib.sha256(data).hexdigest(), os.path.join(self.dir, "f"), opener_for(data))
        self.assertEqual(self.read(path), data)

    def test_hash_mismatch_removes_file(self):
        path = os.path.join(self.dir, "f")
        with self.assertRaises(inst.InstallError):
            inst.download("https://x", "0" * 64, path, opener_for(b"evil"))
        self.assertFalse(os.path.exists(path))

    def test_network_error(self):
        def broken(request, timeout=None):
            raise OSError("нет сети")

        with self.assertRaises(inst.InstallError):
            inst.download("https://x", "0" * 64, os.path.join(self.dir, "f"), broken)

    def test_pinned_versions(self):
        self.assertIn("/v1.10/", inst.RNNOISE_ZIP_URL)
        self.assertIn("/1.4.2/", inst.EQAPO_URL)
        self.assertEqual(len(inst.RNNOISE_ZIP_SHA256), 64)
        self.assertEqual(len(inst.EQAPO_SHA256), 64)


class RnnoiseTests(Base):
    def test_extract_takes_vst_not_lv2(self):
        zip_path = os.path.join(self.dir, "r.zip")
        with open(zip_path, "wb") as f:
            f.write(rnnoise_zip())
        out = inst.extract_rnnoise(zip_path, os.path.join(self.dir, "out"))
        self.assertEqual(self.read(os.path.join(out, "rnnoise_stereo.dll")), b"VST-STEREO")
        self.assertEqual(self.read(os.path.join(out, "rnnoise_mono.dll")), b"VST-MONO")

    def test_extract_missing_member(self):
        zip_path = os.path.join(self.dir, "r.zip")
        with zipfile.ZipFile(zip_path, "w") as z:
            z.writestr("other.txt", b"")
        with self.assertRaises(inst.InstallError):
            inst.extract_rnnoise(zip_path, os.path.join(self.dir, "out"))

    def test_target_dir(self):
        self.assertEqual(
            inst.rnnoise_target_dir(self.config_dir + os.sep),
            os.path.join(self.dir, "EqualizerAPO", "VSTPlugins", "RNNoise"),
        )

    def put_bundled(self):
        folder = os.path.join(self.app_dir, inst.BUNDLED_RNNOISE_DIR)
        os.makedirs(folder)
        for name in inst.RNNOISE_ZIP_MEMBERS:
            with open(os.path.join(folder, name), "wb") as f:
                f.write(b"BUNDLED-" + name.encode())
        return folder

    def test_install_from_bundled_without_network(self):
        self.put_bundled()

        def no_network(request, timeout=None):
            raise AssertionError("сеть не нужна")

        path = inst.install_rnnoise(self.app_dir, self.config_dir, opener=no_network)
        self.assertEqual(path, os.path.join(inst.rnnoise_target_dir(self.config_dir), "rnnoise_stereo.dll"))
        self.assertEqual(self.read(path), b"BUNDLED-rnnoise_stereo.dll")

    def test_install_downloads_when_not_bundled(self):
        data = rnnoise_zip()
        calls = []
        with mock.patch.object(inst, "RNNOISE_ZIP_SHA256", hashlib.sha256(data).hexdigest()):
            path = inst.install_rnnoise(self.app_dir, self.config_dir, opener=opener_for(data, calls))
        self.assertEqual(calls, [inst.RNNOISE_ZIP_URL])
        self.assertEqual(self.read(path), b"VST-STEREO")

    def test_install_requires_equalizer_apo(self):
        with self.assertRaises(inst.InstallError):
            inst.install_rnnoise(self.app_dir, None)

    def test_permission_error_goes_through_uac(self):
        self.put_bundled()
        runs = []

        def fake_run(cmd, **kwargs):
            runs.append(cmd)
            dest = inst.rnnoise_target_dir(self.config_dir)
            os.makedirs(dest, exist_ok=True)
            for name in inst.RNNOISE_ZIP_MEMBERS:
                open(os.path.join(dest, name), "w").close()

        with mock.patch.object(inst, "copy_files", side_effect=PermissionError("Access is denied")):
            path = inst.install_rnnoise(self.app_dir, self.config_dir, run=fake_run)
        self.assertEqual(len(runs), 1)
        self.assertTrue(os.path.isfile(path))

    def test_uac_declined(self):
        self.put_bundled()
        with mock.patch.object(inst, "copy_files", side_effect=PermissionError("Access is denied")):
            with self.assertRaises(inst.InstallError):
                inst.install_rnnoise(self.app_dir, self.config_dir, run=lambda cmd, **kw: None)


class ElevatedCommandTests(unittest.TestCase):
    def test_command_runs_elevated_and_waits(self):
        cmd = inst.elevated_copy_command(r"C:\Temp\src", r"C:\Program Files\EqualizerAPO\VSTPlugins\RNNoise", ["a.dll"])
        self.assertEqual(cmd[:4], ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command"])
        self.assertIn("-Verb RunAs -Wait", cmd[4])

    def test_encoded_script(self):
        cmd = inst.elevated_copy_command(r"C:\Users\O'Neil\src", r"C:\Program Files\EqualizerAPO\VSTPlugins\RNNoise", ["a.dll", "b.dll"])
        encoded = cmd[4].split("'-EncodedCommand','")[1].rstrip("'")
        script = base64.b64decode(encoded).decode("utf-16-le")
        self.assertIn("$ErrorActionPreference = 'Stop'", script)
        self.assertIn("New-Item -ItemType Directory -Force -Path 'C:\\Program Files\\EqualizerAPO\\VSTPlugins\\RNNoise'", script)
        self.assertIn("'C:\\Users\\O''Neil\\src\\a.dll','C:\\Users\\O''Neil\\src\\b.dll'", script.replace("/", "\\"))


class EqualizerApoTests(Base):
    def test_download_installer(self):
        data = b"MZ installer"
        with mock.patch.object(inst, "EQAPO_SHA256", hashlib.sha256(data).hexdigest()):
            path = inst.download_equalizer_apo(self.dir, opener_for(data))
        self.assertTrue(path.endswith("EqualizerAPO-x64-1.4.2.exe"))
        self.assertEqual(self.read(path), data)


if __name__ == "__main__":
    unittest.main()
