"""Core Audio через pycaw/comtypes: без внешних программ.

Переключение по умолчанию идёт через IPolicyConfig. Интерфейс недокументирован,
но им пользуются все подобные утилиты (и SoundVolumeView тоже).

COM инициализируется в каждом потоке при первом обращении, COM-объекты между
потоками не передаются: перечислитель создаётся на каждый вызов.
"""
import threading

import comtypes
from comtypes import GUID
from comtypes.automation import VT_LPWSTR
from pycaw.api.mmdeviceapi import IMMDeviceEnumerator
from pycaw.api.mmdeviceapi.depend.structures import PROPERTYKEY
from pycaw.api.policyconfig import IPolicyConfig
from pycaw.constants import CLSID_CPolicyConfigClient, CLSID_MMDeviceEnumerator

from . import CAPTURE, RENDER, Device, sort_devices

FLOW_IDS = {RENDER: 0, CAPTURE: 1}
DEVICE_STATE_ACTIVE = 0x1
STGM_READ = 0


def _property_key(fmtid, pid):
    key = PROPERTYKEY()
    key.fmtid = GUID(fmtid)
    key.pid = pid
    return key


# «Микрофон»
PKEY_DEVICE_DESC = _property_key("{a45c254e-df1c-4efd-8020-67d146a850e0}", 2)
# «Микрофон (Realtek(R) Audio)»
PKEY_DEVICE_FRIENDLY_NAME = _property_key("{a45c254e-df1c-4efd-8020-67d146a850e0}", 14)
# «Realtek(R) Audio»
PKEY_INTERFACE_FRIENDLY_NAME = _property_key("{026e516e-b814-414b-83cd-856d6fef4822}", 2)

COM_ERRORS = (comtypes.COMError, OSError)


class CoreAudioBackend:
    name = "Core Audio"
    watchable = True

    def __init__(self):
        self._local = threading.local()
        # Проверяем сразу: если интерфейсы не создаются, open_backend уйдёт на запасной вариант.
        self._enumerator()
        self._policy_config()

    def _ensure_com(self):
        if getattr(self._local, "com_ready", False):
            return
        try:
            comtypes.CoInitialize()
        except OSError:
            pass  # поток уже инициализирован в другом режиме COM, работать можно
        self._local.com_ready = True

    def _enumerator(self):
        self._ensure_com()
        return comtypes.CoCreateInstance(
            CLSID_MMDeviceEnumerator, IMMDeviceEnumerator, comtypes.CLSCTX_INPROC_SERVER
        )

    def _policy_config(self):
        self._ensure_com()
        return comtypes.CoCreateInstance(
            CLSID_CPolicyConfigClient, IPolicyConfig, comtypes.CLSCTX_ALL
        )

    @staticmethod
    def _read_string(store, key):
        try:
            value = store.GetValue(key)
        except COM_ERRORS:
            return ""
        try:
            if value.vt != VT_LPWSTR:
                return ""
            return (value.GetValue() or "").strip()
        finally:
            value.clear()

    def _make_device(self, mm_device, flow):
        dev_id = mm_device.GetId()
        try:
            store = mm_device.OpenPropertyStore(STGM_READ)
        except COM_ERRORS:
            return Device(id=dev_id, name=dev_id, flow=flow)

        name = self._read_string(store, PKEY_DEVICE_DESC)
        if not name:
            name = self._read_string(store, PKEY_DEVICE_FRIENDLY_NAME)
        interface = self._read_string(store, PKEY_INTERFACE_FRIENDLY_NAME)
        return Device(id=dev_id, name=name or dev_id, interface=interface, flow=flow)

    def list_devices(self, flow):
        collection = self._enumerator().EnumAudioEndpoints(FLOW_IDS[flow], DEVICE_STATE_ACTIVE)
        return sort_devices(
            self._make_device(collection.Item(i), flow) for i in range(collection.GetCount())
        )

    def get_default_id(self, flow, role):
        try:
            return self._enumerator().GetDefaultAudioEndpoint(FLOW_IDS[flow], role).GetId()
        except COM_ERRORS:
            return None  # нет ни одного активного устройства

    def set_default(self, device_id, roles):
        policy = self._policy_config()
        for role in roles:
            policy.SetDefaultEndpoint(device_id, role)
