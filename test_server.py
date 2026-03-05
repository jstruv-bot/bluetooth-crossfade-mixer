"""
Unit tests for the Bluetooth Crossfade Mixer backend.

Mocks pycaw/comtypes so tests run on any platform (not just Windows).
"""

import json
import unittest
from unittest.mock import patch, MagicMock

# Mock comtypes and pycaw before importing server
import sys

mock_comtypes = MagicMock()
mock_pycaw = MagicMock()
mock_pycaw_constants = MagicMock()
mock_pycaw_constants.EDataFlow.eRender.value = 0
mock_pycaw_constants.DEVICE_STATE.ACTIVE.value = 1

sys.modules["comtypes"] = mock_comtypes
sys.modules["pycaw"] = MagicMock()
sys.modules["pycaw.pycaw"] = mock_pycaw
sys.modules["pycaw.constants"] = mock_pycaw_constants

import server


def _make_mock_device(device_id, name, volume=0.5):
    """Create a mock audio device."""
    dev = MagicMock()
    dev.id = device_id
    dev.FriendlyName = name
    endpoint = MagicMock()
    endpoint.GetMasterVolumeLevelScalar.return_value = volume
    dev.EndpointVolume = endpoint
    return dev


class TestAPIRoutes(unittest.TestCase):
    def setUp(self):
        server.app.testing = True
        self.client = server.app.test_client()

    @patch.object(server, "get_playback_devices")
    def test_get_devices(self, mock_get):
        mock_get.return_value = [
            {"id": "dev1", "name": "Speaker A", "volume": 0.75},
        ]
        resp = self.client.get("/api/devices")
        self.assertEqual(resp.status_code, 200)
        data = json.loads(resp.data)
        self.assertEqual(len(data), 1)
        self.assertEqual(data[0]["name"], "Speaker A")

    @patch.object(server, "set_device_volume")
    def test_set_volume_success(self, mock_set):
        mock_set.return_value = True
        resp = self.client.post(
            "/api/volume",
            data=json.dumps({"device_id": "dev1", "volume": 0.5}),
            content_type="application/json",
        )
        self.assertEqual(resp.status_code, 200)
        data = json.loads(resp.data)
        self.assertTrue(data["success"])

    def test_set_volume_missing_body(self):
        resp = self.client.post("/api/volume", content_type="application/json")
        self.assertEqual(resp.status_code, 400)

    def test_set_volume_missing_device_id(self):
        resp = self.client.post(
            "/api/volume",
            data=json.dumps({"volume": 0.5}),
            content_type="application/json",
        )
        self.assertEqual(resp.status_code, 400)

    def test_set_volume_missing_volume(self):
        resp = self.client.post(
            "/api/volume",
            data=json.dumps({"device_id": "dev1"}),
            content_type="application/json",
        )
        self.assertEqual(resp.status_code, 400)

    def test_set_volume_invalid_volume(self):
        resp = self.client.post(
            "/api/volume",
            data=json.dumps({"device_id": "dev1", "volume": "not_a_number"}),
            content_type="application/json",
        )
        self.assertEqual(resp.status_code, 400)

    def test_set_volume_invalid_device_id_type(self):
        resp = self.client.post(
            "/api/volume",
            data=json.dumps({"device_id": 12345, "volume": 0.5}),
            content_type="application/json",
        )
        self.assertEqual(resp.status_code, 400)

    def test_set_volume_device_id_too_long(self):
        resp = self.client.post(
            "/api/volume",
            data=json.dumps({"device_id": "x" * 600, "volume": 0.5}),
            content_type="application/json",
        )
        self.assertEqual(resp.status_code, 400)

    @patch.object(server, "set_volumes_batch")
    def test_batch_volume(self, mock_batch):
        mock_batch.return_value = {"dev1": True, "dev2": True}
        resp = self.client.post(
            "/api/volume/batch",
            data=json.dumps({"volumes": [
                {"device_id": "dev1", "volume": 0.8},
                {"device_id": "dev2", "volume": 0.3},
            ]}),
            content_type="application/json",
        )
        self.assertEqual(resp.status_code, 200)
        data = json.loads(resp.data)
        self.assertTrue(data["success"])

    def test_batch_volume_missing_body(self):
        resp = self.client.post("/api/volume/batch", content_type="application/json")
        self.assertEqual(resp.status_code, 400)

    @patch.object(server, "get_playback_devices")
    def test_refresh(self, mock_get):
        mock_get.return_value = []
        resp = self.client.post("/api/refresh")
        self.assertEqual(resp.status_code, 200)

    def test_index_page(self):
        resp = self.client.get("/")
        self.assertEqual(resp.status_code, 200)
        self.assertIn(b"Bluetooth Crossfade Mixer", resp.data)


class TestSetDeviceVolume(unittest.TestCase):
    @patch.object(server, "_get_active_render_devices")
    def test_clamps_volume(self, mock_enum):
        dev = _make_mock_device("dev1", "Speaker", 0.5)
        mock_enum.return_value = [dev]
        server.set_device_volume("dev1", 1.5)
        dev.EndpointVolume.SetMasterVolumeLevelScalar.assert_called_with(1.0, None)

    @patch.object(server, "_get_active_render_devices")
    def test_clamps_volume_negative(self, mock_enum):
        dev = _make_mock_device("dev1", "Speaker", 0.5)
        mock_enum.return_value = [dev]
        server.set_device_volume("dev1", -0.5)
        dev.EndpointVolume.SetMasterVolumeLevelScalar.assert_called_with(0.0, None)

    @patch.object(server, "_get_active_render_devices")
    def test_device_not_found(self, mock_enum):
        mock_enum.return_value = []
        result = server.set_device_volume("nonexistent", 0.5)
        self.assertFalse(result)


class TestSetVolumesBatch(unittest.TestCase):
    @patch.object(server, "_get_active_render_devices")
    def test_batch_sets_multiple(self, mock_enum):
        dev1 = _make_mock_device("dev1", "A")
        dev2 = _make_mock_device("dev2", "B")
        mock_enum.return_value = [dev1, dev2]

        results = server.set_volumes_batch([
            {"device_id": "dev1", "volume": 0.8},
            {"device_id": "dev2", "volume": 0.3},
        ])

        self.assertTrue(results["dev1"])
        self.assertTrue(results["dev2"])
        dev1.EndpointVolume.SetMasterVolumeLevelScalar.assert_called_with(0.8, None)
        dev2.EndpointVolume.SetMasterVolumeLevelScalar.assert_called_with(0.3, None)

    @patch.object(server, "_get_active_render_devices")
    def test_batch_missing_device(self, mock_enum):
        mock_enum.return_value = []
        results = server.set_volumes_batch([{"device_id": "dev1", "volume": 0.5}])
        self.assertFalse(results["dev1"])


if __name__ == "__main__":
    unittest.main()
