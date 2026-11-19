"""CoreAudioBackend на фейковых comtypes/pycaw: проверяется логика поверх COM,
а не сам Windows. Настоящий COM - только ручная проверка на Windows."""
import importlib
import sys
import threading
import types
import unittest
from unittest import mock

from tests import stubs  # noqa: F401

import audio_devices as ad

VT_LPWSTR = 31
DESC = ("{a45c254e-df1c-4efd-8020-67d146a850e0}", 2)
FRIENDLY = ("{a45c254e-df1c-4efd-8020-67d146a850e0}", 14)
INTERFACE = ("{026e516e-b814-414b-83cd-856d6fef4822}", 2)


class COMError(Exception):
    pass


class PROPERTYKEY:
    fmtid = None
    pid = None


class PropVariant:
    def __init__(self, value, vt=VT_LPWSTR):
        self.value = value
        self.vt = vt
        self.cleared = False

    def GetValue(self):
        return self.value

    def clear(self):
        self.cleared = True


class Store:
    def __init__(self, props):
        self.props = props
        self.values = []

    def GetValue(self, key):
        found = self.props.get((str(key.fmtid).lower(), key.pid))
        if found is None:
            raise COMError("no property")
        value = found if isinstance(found, PropVariant) else PropVariant(found)
        self.values.append(value)
        return value


class MMDevice:
    def __init__(self, dev_id, props, broken_store=False):
        self.dev_id = dev_id
        self.store = Store(props)
        self.broken_store = broken_store

    def GetId(self):
        return self.dev_id

    def OpenPropertyStore(self, mode):
        if self.broken_store:
            raise COMError("access denied")
        return self.store


class Collection:
    def __init__(self, devices):
        self.devices = devices

    def GetCount(self):
        return len(self.devices)

    def Item(self, i):
        return self.devices[i]


class World:
    def __init__(self):
        self.devices = {0: [], 1: []}
        self.defaults = {}
        self.set_calls = []
        self.co_init = 0
        self.fail_policy = False
        self.enum_args = []


class Enumerator:
    def __init__(self, world):
        self.world = world

    def EnumAudioEndpoints(self, flow, state):
        self.world.enum_args.append((flow, state))
        return Collection(self.world.devices[flow])

    def GetDefaultAudioEndpoint(self, flow, role):
        dev_id = self.world.defaults.get((flow, role))
        if dev_id is None:
            raise COMError("E_NOTFOUND")
        return MMDevice(dev_id, {})


class Policy:
    def __init__(self, world):
        self.world = world

    def SetDefaultEndpoint(self, dev_id, role):
        self.world.set_calls.append((dev_id, role))


def fake_modules(world):
    IMMDeviceEnumerator = type("IMMDeviceEnumerator", (), {})
    IPolicyConfig = type("IPolicyConfig", (), {})

    comtypes = types.ModuleType("comtypes")
    comtypes.COMError = COMError
    comtypes.GUID = lambda s: s
    comtypes.CLSCTX_INPROC_SERVER = 1
    comtypes.CLSCTX_ALL = 23

    def co_initialize():
        world.co_init += 1

    def co_create_instance(clsid, interface, ctx):
        if interface is IMMDeviceEnumerator:
            return Enumerator(world)
        if world.fail_policy:
            raise OSError("class not registered")
        return Policy(world)

    comtypes.CoInitialize = co_initialize
    comtypes.CoCreateInstance = co_create_instance
    automation = types.ModuleType("comtypes.automation")
    automation.VT_LPWSTR = VT_LPWSTR

    mmdeviceapi = types.ModuleType("pycaw.api.mmdeviceapi")
    mmdeviceapi.IMMDeviceEnumerator = IMMDeviceEnumerator
    structures = types.ModuleType("pycaw.api.mmdeviceapi.depend.structures")
    structures.PROPERTYKEY = PROPERTYKEY
    policyconfig = types.ModuleType("pycaw.api.policyconfig")
    policyconfig.IPolicyConfig = IPolicyConfig
    constants = types.ModuleType("pycaw.constants")
    constants.CLSID_CPolicyConfigClient = "policy"
    constants.CLSID_MMDeviceEnumerator = "enumerator"

    return {
        "comtypes": comtypes,
        "comtypes.automation": automation,
        "pycaw": types.ModuleType("pycaw"),
        "pycaw.api": types.ModuleType("pycaw.api"),
        "pycaw.api.mmdeviceapi": mmdeviceapi,
        "pycaw.api.mmdeviceapi.depend": types.ModuleType("pycaw.api.mmdeviceapi.depend"),
        "pycaw.api.mmdeviceapi.depend.structures": structures,
        "pycaw.api.policyconfig": policyconfig,
        "pycaw.constants": constants,
    }


def mic(dev_id, name=None, interface=None, friendly=None, **kw):
    props = {}
    if name is not None:
        props[DESC] = name
    if interface is not None:
        props[INTERFACE] = interface
    if friendly is not None:
        props[FRIENDLY] = friendly
    return MMDevice(dev_id, props, **kw)


class CoreAudioTests(unittest.TestCase):
    def setUp(self):
        self.world = World()
        patcher = mock.patch.dict(sys.modules, fake_modules(self.world))
        patcher.start()
        self.addCleanup(patcher.stop)
        sys.modules.pop("audio_devices.core", None)
        self.addCleanup(sys.modules.pop, "audio_devices.core", None)
        self.core = importlib.import_module("audio_devices.core")

    def backend(self):
        return self.core.CoreAudioBackend()

    def test_lists_active_capture_devices_sorted(self):
        self.world.devices[1] = [
            mic("{b}", "Микрофон", "Realtek(R) Audio"),
            mic("{a}", "Гарнитура", "Jabra"),
        ]
        devices = self.backend().list_devices(ad.CAPTURE)
        self.assertEqual(
            devices,
            [
                ad.Device("{a}", "Гарнитура", "Jabra", ad.CAPTURE),
                ad.Device("{b}", "Микрофон", "Realtek(R) Audio", ad.CAPTURE),
            ],
        )
        self.assertEqual(self.world.enum_args[-1], (1, 0x1))

    def test_render_flow_id(self):
        self.backend().list_devices(ad.RENDER)
        self.assertEqual(self.world.enum_args[-1], (0, 0x1))

    def test_name_fallbacks(self):
        self.world.devices[1] = [
            mic("{f}", friendly="Микрофон (USB)"),
            mic("{n}"),
            mic("{x}", broken_store=True),
            mic("{w}", name=PropVariant(None, vt=19), friendly="Другой"),
        ]
        names = {d.id: d.name for d in self.backend().list_devices(ad.CAPTURE)}
        self.assertEqual(names, {"{f}": "Микрофон (USB)", "{n}": "{n}", "{x}": "{x}", "{w}": "Другой"})

    def test_property_values_are_freed(self):
        device = mic("{a}", "Гарнитура", "Jabra")
        self.world.devices[1] = [device]
        self.backend().list_devices(ad.CAPTURE)
        self.assertTrue(device.store.values)
        self.assertTrue(all(v.cleared for v in device.store.values))

    def test_get_default_id(self):
        self.world.defaults[(1, ad.ROLE_CONSOLE)] = "{a}"
        backend = self.backend()
        self.assertEqual(backend.get_default_id(ad.CAPTURE, ad.ROLE_CONSOLE), "{a}")
        self.assertIsNone(backend.get_default_id(ad.CAPTURE, ad.ROLE_COMMUNICATIONS))

    def test_set_default_for_each_role(self):
        self.backend().set_default("{a}", ad.ALL_ROLES)
        self.assertEqual(self.world.set_calls, [("{a}", 0), ("{a}", 1), ("{a}", 2)])

    def test_init_fails_without_policy_config(self):
        self.world.fail_policy = True
        with self.assertRaises(OSError):
            self.backend()

    def test_com_initialized_once_per_thread(self):
        backend = self.backend()
        backend.list_devices(ad.CAPTURE)
        backend.get_default_id(ad.CAPTURE, 0)
        self.assertEqual(self.world.co_init, 1)

        thread = threading.Thread(target=backend.list_devices, args=(ad.CAPTURE,))
        thread.start()
        thread.join()
        self.assertEqual(self.world.co_init, 2)

    def test_co_initialize_mode_conflict_is_ignored(self):
        def conflict():
            raise OSError("RPC_E_CHANGED_MODE")

        sys.modules["comtypes"].CoInitialize = conflict
        self.backend().list_devices(ad.CAPTURE)


if __name__ == "__main__":
    unittest.main()
