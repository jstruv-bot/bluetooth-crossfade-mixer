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

# Mock flask-socketio
mock_socketio_module = MagicMock()
mock_socketio_class = MagicMock()
mock_socketio_instance = MagicMock()
mock_socketio_class.return_value = mock_socketio_instance
mock_socketio_module.SocketIO = mock_socketio_class
mock_socketio_module.emit = MagicMock()
sys.modules["flask_socketio"] = mock_socketio_module

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
        # Reset state between tests
        server._muted_devices.clear()
        server._device_groups.clear()
        server._last_volumes.clear()
        server._device_eq.clear()
        server._previous_device_ids.clear()

    @patch.object(server, "get_playback_devices")
    def test_get_devices(self, mock_get):
        mock_get.return_value = [
            {"id": "dev1", "name": "Speaker A", "volume": 0.75,
             "muted": False, "group": None, "eq": {"bass": 0.0, "treble": 0.0}},
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
        self.assertTrue(json.loads(resp.data)["success"])

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
        self.assertTrue(json.loads(resp.data)["success"])

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
    def setUp(self):
        server._last_volumes.clear()

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

    @patch.object(server, "_get_active_render_devices")
    def test_remembers_volume_for_reconnect(self, mock_enum):
        dev = _make_mock_device("dev1", "Speaker", 0.5)
        mock_enum.return_value = [dev]
        server.set_device_volume("dev1", 0.75)
        self.assertEqual(server._last_volumes["dev1"], 0.75)


class TestSetVolumesBatch(unittest.TestCase):
    def setUp(self):
        server._last_volumes.clear()

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

    @patch.object(server, "_get_active_render_devices")
    def test_batch_remembers_volumes(self, mock_enum):
        dev = _make_mock_device("dev1", "A")
        mock_enum.return_value = [dev]
        server.set_volumes_batch([{"device_id": "dev1", "volume": 0.6}])
        self.assertEqual(server._last_volumes["dev1"], 0.6)


class TestMuteAPI(unittest.TestCase):
    def setUp(self):
        server.app.testing = True
        self.client = server.app.test_client()
        server._muted_devices.clear()

    @patch.object(server, "set_device_volume")
    def test_mute_device(self, mock_set):
        mock_set.return_value = True
        resp = self.client.post(
            "/api/mute",
            data=json.dumps({"device_id": "dev1", "muted": True}),
            content_type="application/json",
        )
        self.assertEqual(resp.status_code, 200)
        data = json.loads(resp.data)
        self.assertTrue(data["success"])
        self.assertTrue(data["muted"])
        self.assertIn("dev1", server._muted_devices)

    @patch.object(server, "set_device_volume")
    def test_unmute_device(self, mock_set):
        server._muted_devices.add("dev1")
        resp = self.client.post(
            "/api/mute",
            data=json.dumps({"device_id": "dev1", "muted": False}),
            content_type="application/json",
        )
        self.assertEqual(resp.status_code, 200)
        self.assertNotIn("dev1", server._muted_devices)

    def test_mute_invalid_body(self):
        resp = self.client.post("/api/mute", content_type="application/json")
        self.assertEqual(resp.status_code, 400)

    def test_mute_invalid_muted_type(self):
        resp = self.client.post(
            "/api/mute",
            data=json.dumps({"device_id": "dev1", "muted": "yes"}),
            content_type="application/json",
        )
        self.assertEqual(resp.status_code, 400)


class TestGroupsAPI(unittest.TestCase):
    def setUp(self):
        server.app.testing = True
        self.client = server.app.test_client()
        server._device_groups.clear()

    def test_create_group(self):
        resp = self.client.post(
            "/api/groups",
            data=json.dumps({"name": "Living Room", "device_ids": ["dev1", "dev2"]}),
            content_type="application/json",
        )
        self.assertEqual(resp.status_code, 200)
        data = json.loads(resp.data)
        self.assertTrue(data["success"])
        self.assertIn("Living Room", data["groups"])

    def test_list_groups(self):
        server._device_groups["Zone A"] = ["dev1"]
        resp = self.client.get("/api/groups")
        self.assertEqual(resp.status_code, 200)
        data = json.loads(resp.data)
        self.assertIn("Zone A", data)

    def test_delete_group(self):
        server._device_groups["Zone A"] = ["dev1"]
        resp = self.client.delete("/api/groups/Zone%20A")
        self.assertEqual(resp.status_code, 200)
        self.assertNotIn("Zone A", server._device_groups)

    def test_create_group_invalid_name(self):
        resp = self.client.post(
            "/api/groups",
            data=json.dumps({"name": "", "device_ids": ["dev1"]}),
            content_type="application/json",
        )
        self.assertEqual(resp.status_code, 400)

    def test_create_group_name_too_long(self):
        resp = self.client.post(
            "/api/groups",
            data=json.dumps({"name": "x" * 100, "device_ids": ["dev1"]}),
            content_type="application/json",
        )
        self.assertEqual(resp.status_code, 400)

    def test_create_group_moves_devices_from_old_group(self):
        server._device_groups["Old"] = ["dev1", "dev2"]
        self.client.post(
            "/api/groups",
            data=json.dumps({"name": "New", "device_ids": ["dev1"]}),
            content_type="application/json",
        )
        # dev1 should be removed from "Old" and added to "New"
        self.assertNotIn("dev1", server._device_groups.get("Old", []))
        self.assertIn("dev1", server._device_groups["New"])


class TestEQAPI(unittest.TestCase):
    def setUp(self):
        server.app.testing = True
        self.client = server.app.test_client()
        server._device_eq.clear()

    def test_set_eq(self):
        resp = self.client.post(
            "/api/eq",
            data=json.dumps({"device_id": "dev1", "bass": 0.5, "treble": -0.3}),
            content_type="application/json",
        )
        self.assertEqual(resp.status_code, 200)
        data = json.loads(resp.data)
        self.assertTrue(data["success"])
        self.assertEqual(data["eq"]["bass"], 0.5)
        self.assertEqual(data["eq"]["treble"], -0.3)

    def test_get_eq_default(self):
        resp = self.client.get("/api/eq/dev1")
        self.assertEqual(resp.status_code, 200)
        data = json.loads(resp.data)
        self.assertEqual(data["bass"], 0.0)
        self.assertEqual(data["treble"], 0.0)

    def test_set_eq_clamps(self):
        resp = self.client.post(
            "/api/eq",
            data=json.dumps({"device_id": "dev1", "bass": 5.0, "treble": -5.0}),
            content_type="application/json",
        )
        data = json.loads(resp.data)
        self.assertEqual(data["eq"]["bass"], 1.0)
        self.assertEqual(data["eq"]["treble"], -1.0)

    def test_set_eq_invalid_body(self):
        resp = self.client.post("/api/eq", content_type="application/json")
        self.assertEqual(resp.status_code, 400)

    def test_set_eq_invalid_device_id(self):
        resp = self.client.post(
            "/api/eq",
            data=json.dumps({"device_id": 123, "bass": 0.0}),
            content_type="application/json",
        )
        self.assertEqual(resp.status_code, 400)


class TestAutoReconnect(unittest.TestCase):
    def setUp(self):
        server._last_volumes.clear()
        server._previous_device_ids.clear()

    @patch.object(server, "socketio")
    def test_reconnect_restores_volume(self, mock_sio):
        # Simulate: dev1 was known with volume 0.8, then disappeared, now reappears
        server._last_volumes["dev1"] = 0.8
        server._previous_device_ids.update(set())  # empty = first scan

        dev1 = _make_mock_device("dev1", "Speaker A", 0.5)
        current_ids = {"dev1"}

        # First call to establish baseline
        server._handle_reconnections(current_ids, [dev1])

        # Now simulate dev1 disappearing and reappearing
        server._previous_device_ids.clear()  # it disappeared
        server._handle_reconnections(current_ids, [dev1])

        # Volume should have been restored
        dev1.EndpointVolume.SetMasterVolumeLevelScalar.assert_called_with(0.8, None)


class TestValidateDeviceId(unittest.TestCase):
    def test_valid(self):
        self.assertTrue(server._validate_device_id("abc123"))

    def test_too_long(self):
        self.assertFalse(server._validate_device_id("x" * 600))

    def test_not_string(self):
        self.assertFalse(server._validate_device_id(123))

    def test_empty(self):
        self.assertFalse(server._validate_device_id(""))

    def test_none(self):
        self.assertFalse(server._validate_device_id(None))


class TestSpotifyAPI(unittest.TestCase):
    def setUp(self):
        server.app.testing = True
        server.app.config["SECRET_KEY"] = "test-secret"
        self.client = server.app.test_client()
        # Reset Spotify state
        with server._spotify_lock:
            server._spotify_tokens["access_token"] = None
            server._spotify_tokens["refresh_token"] = None
            server._spotify_tokens["expires_at"] = 0
            server._spotify_tokens["client_id"] = ""

    def test_spotify_status_disconnected(self):
        resp = self.client.get("/api/spotify/status")
        self.assertEqual(resp.status_code, 200)
        data = json.loads(resp.data)
        self.assertFalse(data["connected"])

    def test_spotify_status_connected(self):
        import time
        with server._spotify_lock:
            server._spotify_tokens["access_token"] = "test_token"
            server._spotify_tokens["expires_at"] = time.time() + 3600
        resp = self.client.get("/api/spotify/status")
        data = json.loads(resp.data)
        self.assertTrue(data["connected"])

    def test_spotify_now_playing_default(self):
        resp = self.client.get("/api/spotify/now-playing")
        self.assertEqual(resp.status_code, 200)
        data = json.loads(resp.data)
        self.assertFalse(data["is_playing"])

    def test_spotify_play_unauthorized(self):
        resp = self.client.post("/api/spotify/play")
        self.assertEqual(resp.status_code, 401)

    def test_spotify_pause_unauthorized(self):
        resp = self.client.post("/api/spotify/pause")
        self.assertEqual(resp.status_code, 401)

    def test_spotify_next_unauthorized(self):
        resp = self.client.post("/api/spotify/next")
        self.assertEqual(resp.status_code, 401)

    def test_spotify_previous_unauthorized(self):
        resp = self.client.post("/api/spotify/previous")
        self.assertEqual(resp.status_code, 401)

    def test_spotify_disconnect(self):
        import time
        with server._spotify_lock:
            server._spotify_tokens["access_token"] = "test_token"
            server._spotify_tokens["expires_at"] = time.time() + 3600
        resp = self.client.post("/api/spotify/disconnect")
        self.assertEqual(resp.status_code, 200)
        self.assertTrue(json.loads(resp.data)["success"])
        # Verify tokens cleared
        with server._spotify_lock:
            self.assertIsNone(server._spotify_tokens["access_token"])

    def test_spotify_login_no_client_id(self):
        # No client_id and no env var
        resp = self.client.get("/api/spotify/login")
        self.assertEqual(resp.status_code, 400)

    def test_spotify_login_with_client_id(self):
        resp = self.client.get("/api/spotify/login?client_id=test_client_123")
        # Should redirect to Spotify authorize page
        self.assertEqual(resp.status_code, 302)
        self.assertIn("accounts.spotify.com/authorize", resp.headers["Location"])

    def test_spotify_callback_no_code(self):
        resp = self.client.get("/api/spotify/callback")
        self.assertEqual(resp.status_code, 400)

    def test_spotify_callback_error(self):
        resp = self.client.get("/api/spotify/callback?error=access_denied")
        self.assertEqual(resp.status_code, 400)


class TestAudioLevelsAPI(unittest.TestCase):
    def setUp(self):
        server.app.testing = True
        self.client = server.app.test_client()

    def test_get_levels(self):
        with server._levels_lock:
            server._audio_levels.clear()
            server._audio_levels["dev1"] = 0.5
        resp = self.client.get("/api/levels")
        self.assertEqual(resp.status_code, 200)
        data = json.loads(resp.data)
        self.assertEqual(data["dev1"], 0.5)

    def test_get_levels_empty(self):
        with server._levels_lock:
            server._audio_levels.clear()
        resp = self.client.get("/api/levels")
        self.assertEqual(resp.status_code, 200)
        data = json.loads(resp.data)
        self.assertEqual(data, {})


class TestSpotifyHelpers(unittest.TestCase):
    def test_pkce_generation(self):
        verifier, challenge = server._spotify_generate_pkce()
        self.assertIsInstance(verifier, str)
        self.assertIsInstance(challenge, str)
        self.assertGreater(len(verifier), 40)
        self.assertGreater(len(challenge), 20)
        # Challenge should not contain padding
        self.assertNotIn("=", challenge)

    def test_token_valid_when_expired(self):
        with server._spotify_lock:
            server._spotify_tokens["access_token"] = "test"
            server._spotify_tokens["expires_at"] = 0
        self.assertFalse(server._spotify_token_valid())

    def test_token_valid_when_none(self):
        with server._spotify_lock:
            server._spotify_tokens["access_token"] = None
        self.assertFalse(server._spotify_token_valid())

    def test_token_valid_when_fresh(self):
        import time
        with server._spotify_lock:
            server._spotify_tokens["access_token"] = "test"
            server._spotify_tokens["expires_at"] = time.time() + 3600
        self.assertTrue(server._spotify_token_valid())


if __name__ == "__main__":
    unittest.main()
