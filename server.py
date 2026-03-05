"""
Bluetooth Speaker Crossfade Mixer — Backend
Flask API + pycaw device enumeration for per-device volume control.
Features: crossfade curves, mute, presets, groups, WebSocket, EQ,
          auto-reconnect, Spotify integration, audio level metering.
"""

import html
import logging
import webbrowser
import threading
import os
import time
import hashlib
import secrets
import base64
import urllib.parse
from flask import Flask, render_template, jsonify, request, redirect, session
from flask_socketio import SocketIO, emit
import requests as http_requests

# pycaw / COM imports for Windows Core Audio
import comtypes
from pycaw.pycaw import AudioUtilities
from pycaw.constants import EDataFlow, DEVICE_STATE

# IAudioMeterInformation for audio level metering
try:
    from comtypes import GUID
    _IID_IAudioMeterInformation = GUID("{C02216F6-8C67-4B5B-9D00-D008E73E0064}")
    _METER_AVAILABLE = True
except Exception:
    _METER_AVAILABLE = False

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Flask app setup
# ---------------------------------------------------------------------------

app = Flask(__name__, template_folder="templates")
app.config["SECRET_KEY"] = os.urandom(24)
socketio = SocketIO(app, cors_allowed_origins=None, async_mode="threading")

# Thread-local storage for COM initialization tracking
_com_initialized = threading.local()

# ---------------------------------------------------------------------------
# Persistent state (in-memory, survives across requests)
# ---------------------------------------------------------------------------

# Maximum length for device_id strings to prevent abuse
_MAX_DEVICE_ID_LEN = 512
_MAX_GROUP_NAME_LEN = 64
_MAX_GROUPS = 20

# Muted devices: set of device_id
_muted_devices = set()
_muted_lock = threading.Lock()

# Device groups: {group_name: [device_id, ...]}
_device_groups = {}
_groups_lock = threading.Lock()

# Last known volumes: {device_id: float} — for auto-reconnect restore
_last_volumes = {}
_volumes_lock = threading.Lock()

# Per-device EQ settings: {device_id: {"bass": float, "treble": float}}
_device_eq = {}
_eq_lock = threading.Lock()

# Previous device set for detecting reconnections
_previous_device_ids = set()
_prev_lock = threading.Lock()

# Audio level cache: {device_id: float (0.0-1.0)}
_audio_levels = {}
_levels_lock = threading.Lock()

# Cached device objects for audio metering (avoids re-enumerating COM at 10fps)
_cached_meter_devices = []
_cached_meter_ts = 0.0
_METER_CACHE_TTL = 3.0  # reuse device list for 3 seconds

# ---------------------------------------------------------------------------
# Spotify state
# ---------------------------------------------------------------------------

_SPOTIFY_CLIENT_ID = os.environ.get("SPOTIFY_CLIENT_ID", "")
_SPOTIFY_REDIRECT_URI = "http://127.0.0.1:5123/api/spotify/callback"
_SPOTIFY_SCOPES = "user-read-playback-state user-modify-playback-state user-read-currently-playing"

_spotify_lock = threading.Lock()
_spotify_tokens = {
    "access_token": None,
    "refresh_token": None,
    "expires_at": 0,
    "client_id": "",
}
_spotify_now_playing = {
    "is_playing": False,
    "track": None,
    "artist": None,
    "album": None,
    "album_art": None,
    "progress_ms": 0,
    "duration_ms": 0,
}

# ---------------------------------------------------------------------------
# Device enumeration helpers
# ---------------------------------------------------------------------------


def _init_com():
    """Initialize COM for the current thread (once per thread)."""
    if not getattr(_com_initialized, "done", False):
        try:
            comtypes.CoInitialize()
        except OSError:
            pass
        _com_initialized.done = True


def _get_active_render_devices():
    """Return the list of active render (playback) devices."""
    _init_com()
    return AudioUtilities.GetAllDevices(
        data_flow=EDataFlow.eRender.value,
        device_state=DEVICE_STATE.ACTIVE.value,
    )


def get_playback_devices():
    """
    Enumerate all active audio *render* (playback) devices and return a list
    of dicts with id, name, current volume, mute state, group, and EQ.
    Also handles auto-reconnect volume restoration.
    """
    devices_info = []

    try:
        all_devices = _get_active_render_devices()
    except Exception as exc:
        log.error("Failed to list devices: %s", exc)
        return devices_info

    current_ids = set()

    # Snapshot shared state once to avoid per-device lock contention
    with _muted_lock:
        muted_snapshot = set(_muted_devices)
    with _eq_lock:
        eq_snapshot = dict(_device_eq)
    with _groups_lock:
        # Build reverse lookup: device_id -> group_name
        device_to_group = {}
        for gname, members in _device_groups.items():
            for mid in members:
                device_to_group[mid] = gname

    for device in all_devices:
        try:
            friendly_name = device.FriendlyName
            device_id = device.id
            if not friendly_name or not device_id:
                continue

            current_ids.add(device_id)

            volume_level = None
            try:
                endpoint_volume = device.EndpointVolume
                if endpoint_volume is not None:
                    volume_level = endpoint_volume.GetMasterVolumeLevelScalar()
            except Exception:
                pass

            devices_info.append({
                "id": device_id,
                "name": friendly_name,
                "volume": round(volume_level, 4) if volume_level is not None else None,
                "muted": device_id in muted_snapshot,
                "group": device_to_group.get(device_id),
                "eq": eq_snapshot.get(device_id, {"bass": 0.0, "treble": 0.0}),
            })

        except Exception as exc:
            log.warning("Skipping device: %s", exc)
            continue

    # Auto-reconnect: detect newly appeared devices and restore their volume
    _handle_reconnections(current_ids, all_devices)

    return devices_info


def _handle_reconnections(current_ids, all_devices):
    """Restore volume for devices that just reconnected."""
    with _prev_lock:
        previously_known = _previous_device_ids.copy()
        _previous_device_ids.clear()
        _previous_device_ids.update(current_ids)

    newly_connected = current_ids - previously_known
    if not newly_connected:
        return

    with _volumes_lock:
        restore_map = {
            did: _last_volumes[did]
            for did in newly_connected
            if did in _last_volumes
        }

    if not restore_map:
        return

    for device in all_devices:
        try:
            if device.id in restore_map:
                endpoint_volume = device.EndpointVolume
                if endpoint_volume is not None:
                    vol = restore_map[device.id]
                    endpoint_volume.SetMasterVolumeLevelScalar(vol, None)
                    log.info("Auto-restored volume %.2f on reconnected device %s", vol, device.id)
                    socketio.emit("device_reconnected", {
                        "device_id": device.id,
                        "restored_volume": vol,
                    })
        except Exception as exc:
            log.error("Failed to restore volume on %s: %s", device.id, exc)


def set_device_volume(device_id, volume):
    """
    Set the master volume on the device identified by *device_id*.
    Returns True on success, False on failure.
    """
    volume = max(0.0, min(1.0, float(volume)))

    try:
        all_devices = _get_active_render_devices()
    except Exception as exc:
        log.error("Failed to enumerate devices: %s", exc)
        return False

    for device in all_devices:
        try:
            if device.id == device_id:
                endpoint_volume = device.EndpointVolume
                if endpoint_volume is None:
                    log.warning("Device has no volume interface: %s", device_id)
                    return False
                endpoint_volume.SetMasterVolumeLevelScalar(volume, None)
                # Remember volume for auto-reconnect
                with _volumes_lock:
                    _last_volumes[device_id] = volume
                return True
        except Exception as exc:
            log.error("Error setting volume on %s: %s", device_id, exc)
            return False

    log.warning("Device not found: %s", device_id)
    return False


def set_volumes_batch(volumes):
    """
    Set volume on multiple devices in a single enumeration pass.
    Returns a dict mapping device_id -> bool (success/failure).
    """
    results = {}
    lookup = {}
    for entry in volumes:
        did = entry.get("device_id")
        vol = entry.get("volume")
        if did and vol is not None:
            lookup[did] = max(0.0, min(1.0, float(vol)))

    if not lookup:
        return results

    try:
        all_devices = _get_active_render_devices()
    except Exception as exc:
        log.error("Failed to enumerate devices: %s", exc)
        return {did: False for did in lookup}

    successfully_set = {}
    for device in all_devices:
        try:
            if device.id in lookup:
                endpoint_volume = device.EndpointVolume
                if endpoint_volume is None:
                    results[device.id] = False
                else:
                    endpoint_volume.SetMasterVolumeLevelScalar(lookup[device.id], None)
                    results[device.id] = True
                    successfully_set[device.id] = lookup[device.id]
        except Exception as exc:
            log.error("Error setting volume on %s: %s", device.id, exc)
            results[device.id] = False

    # Batch-update remembered volumes for auto-reconnect
    if successfully_set:
        with _volumes_lock:
            _last_volumes.update(successfully_set)

    for did in lookup:
        if did not in results:
            results[did] = False

    return results


def _validate_device_id(device_id):
    """Return True if device_id is a valid string within length limits."""
    return isinstance(device_id, str) and 0 < len(device_id) <= _MAX_DEVICE_ID_LEN


def get_audio_levels():
    """Read peak audio levels for all active render devices via IAudioMeterInformation.

    Caches the device list for _METER_CACHE_TTL seconds to avoid
    re-enumerating COM devices at 10fps.
    """
    global _cached_meter_devices, _cached_meter_ts

    levels = {}
    if not _METER_AVAILABLE:
        return levels

    now = time.monotonic()
    if now - _cached_meter_ts > _METER_CACHE_TTL:
        try:
            _cached_meter_devices = _get_active_render_devices()
            _cached_meter_ts = now
        except Exception:
            return levels

    for device in _cached_meter_devices:
        try:
            endpoint = device._dev
            meter_info = endpoint.Activate(_IID_IAudioMeterInformation, 0, None)
            peak = meter_info.GetPeakValue()
            levels[device.id] = round(max(0.0, min(1.0, peak)), 4)
        except Exception:
            levels[device.id] = 0.0

    with _levels_lock:
        _audio_levels.clear()
        _audio_levels.update(levels)

    return levels


# ---------------------------------------------------------------------------
# Spotify helpers
# ---------------------------------------------------------------------------


def _spotify_generate_pkce():
    """Generate PKCE code verifier and challenge."""
    verifier = secrets.token_urlsafe(64)[:128]
    digest = hashlib.sha256(verifier.encode("ascii")).digest()
    challenge = base64.urlsafe_b64encode(digest).rstrip(b"=").decode("ascii")
    return verifier, challenge


def _spotify_token_valid():
    """Check if we have a valid (non-expired) Spotify access token."""
    with _spotify_lock:
        return (_spotify_tokens["access_token"] is not None
                and time.time() < _spotify_tokens["expires_at"])


def _spotify_refresh():
    """Refresh the Spotify access token using the refresh token."""
    with _spotify_lock:
        refresh_token = _spotify_tokens.get("refresh_token")
        client_id = _spotify_tokens.get("client_id") or _SPOTIFY_CLIENT_ID

    if not refresh_token or not client_id:
        return False

    try:
        resp = http_requests.post("https://accounts.spotify.com/api/token", data={
            "grant_type": "refresh_token",
            "refresh_token": refresh_token,
            "client_id": client_id,
        }, timeout=10)

        if resp.status_code != 200:
            log.warning("Spotify token refresh failed: %s", resp.text)
            return False

        data = resp.json()
        with _spotify_lock:
            _spotify_tokens["access_token"] = data["access_token"]
            _spotify_tokens["expires_at"] = time.time() + data.get("expires_in", 3600) - 60
            if "refresh_token" in data:
                _spotify_tokens["refresh_token"] = data["refresh_token"]
        return True
    except Exception as exc:
        log.error("Spotify refresh error: %s", exc)
        return False


def _spotify_get_headers():
    """Return Authorization header dict, refreshing token if needed."""
    if not _spotify_token_valid():
        _spotify_refresh()
    with _spotify_lock:
        token = _spotify_tokens.get("access_token")
    if not token:
        return None
    return {"Authorization": f"Bearer {token}"}


def _spotify_poll_now_playing():
    """Poll Spotify for the current playback state."""
    headers = _spotify_get_headers()
    if not headers:
        return

    try:
        resp = http_requests.get(
            "https://api.spotify.com/v1/me/player/currently-playing",
            headers=headers, timeout=5)

        if resp.status_code == 204 or resp.status_code == 401:
            with _spotify_lock:
                _spotify_now_playing["is_playing"] = False
                _spotify_now_playing["track"] = None
            return

        if resp.status_code != 200:
            return

        data = resp.json()
        item = data.get("item")
        if not item:
            return

        artists = ", ".join(a["name"] for a in item.get("artists", []))
        images = item.get("album", {}).get("images", [])
        art_url = images[0]["url"] if images else None

        with _spotify_lock:
            _spotify_now_playing.update({
                "is_playing": data.get("is_playing", False),
                "track": item.get("name"),
                "artist": artists,
                "album": item.get("album", {}).get("name"),
                "album_art": art_url,
                "progress_ms": data.get("progress_ms", 0),
                "duration_ms": item.get("duration_ms", 0),
            })
    except Exception as exc:
        log.error("Spotify poll error: %s", exc)


# ---------------------------------------------------------------------------
# Flask routes
# ---------------------------------------------------------------------------


@app.route("/")
def index():
    """Serve the frontend single-page application."""
    return render_template("index.html")


@app.route("/api/devices", methods=["GET"])
def api_devices():
    """Return the current list of active playback devices as JSON."""
    return jsonify(get_playback_devices())


@app.route("/api/volume", methods=["POST"])
def api_volume():
    """Set volume on a specific device."""
    data = request.get_json(silent=True)
    if not data:
        return jsonify({"success": False, "error": "Invalid or missing JSON body"}), 400

    device_id = data.get("device_id")
    volume = data.get("volume")

    if not _validate_device_id(device_id):
        return jsonify({"success": False, "error": "Invalid 'device_id'"}), 400
    if volume is None:
        return jsonify({"success": False, "error": "Missing 'volume'"}), 400

    try:
        volume = float(volume)
    except (TypeError, ValueError):
        return jsonify({"success": False, "error": "'volume' must be a number"}), 400

    ok = set_device_volume(device_id, volume)
    if ok:
        return jsonify({"success": True})
    return jsonify({"success": False, "error": "Failed to set volume"}), 500


@app.route("/api/volume/batch", methods=["POST"])
def api_volume_batch():
    """Set volume on multiple devices in one call."""
    data = request.get_json(silent=True)
    if not data or "volumes" not in data:
        return jsonify({"success": False, "error": "Expected JSON with 'volumes' array"}), 400

    volumes = data["volumes"]
    if not isinstance(volumes, list) or len(volumes) > 20:
        return jsonify({"success": False, "error": "'volumes' must be an array (max 20)"}), 400

    for entry in volumes:
        did = entry.get("device_id")
        vol = entry.get("volume")
        if not _validate_device_id(did):
            return jsonify({"success": False, "error": "Invalid device_id in batch"}), 400
        if vol is None:
            return jsonify({"success": False, "error": "Missing 'volume' in batch entry"}), 400
        try:
            float(vol)
        except (TypeError, ValueError):
            return jsonify({"success": False, "error": "'volume' must be a number"}), 400

    results = set_volumes_batch(volumes)
    all_ok = all(results.values())
    return jsonify({"success": all_ok, "results": results}), (200 if all_ok else 207)


@app.route("/api/refresh", methods=["POST"])
def api_refresh():
    """Re-scan devices and return the updated list."""
    devices = get_playback_devices()
    socketio.emit("devices_updated", devices)
    return jsonify(devices)


# ---------------------------------------------------------------------------
# Mute API
# ---------------------------------------------------------------------------


@app.route("/api/mute", methods=["POST"])
def api_mute():
    """
    Toggle mute on a device.
    Expects JSON: {"device_id": "...", "muted": true/false}
    When muted, the device volume is set to 0 on the system.
    """
    data = request.get_json(silent=True)
    if not data:
        return jsonify({"success": False, "error": "Invalid JSON"}), 400

    device_id = data.get("device_id")
    muted = data.get("muted")

    if not _validate_device_id(device_id):
        return jsonify({"success": False, "error": "Invalid 'device_id'"}), 400
    if not isinstance(muted, bool):
        return jsonify({"success": False, "error": "'muted' must be boolean"}), 400

    with _muted_lock:
        if muted:
            _muted_devices.add(device_id)
        else:
            _muted_devices.discard(device_id)

    # Set system volume to 0 when muting
    if muted:
        set_device_volume(device_id, 0.0)

    socketio.emit("mute_changed", {"device_id": device_id, "muted": muted})
    return jsonify({"success": True, "device_id": device_id, "muted": muted})


# ---------------------------------------------------------------------------
# Device Groups API
# ---------------------------------------------------------------------------


@app.route("/api/groups", methods=["GET"])
def api_groups_list():
    """Return all device groups."""
    with _groups_lock:
        groups_copy = dict(_device_groups)
    return jsonify(groups_copy)


@app.route("/api/groups", methods=["POST"])
def api_groups_create():
    """
    Create or update a device group.
    Expects JSON: {"name": "...", "device_ids": ["...", ...]}
    """
    data = request.get_json(silent=True)
    if not data:
        return jsonify({"success": False, "error": "Invalid JSON"}), 400

    name = data.get("name")
    device_ids = data.get("device_ids")

    if not isinstance(name, str) or not name.strip() or len(name) > _MAX_GROUP_NAME_LEN:
        return jsonify({"success": False, "error": "Invalid group name"}), 400

    if not isinstance(device_ids, list):
        return jsonify({"success": False, "error": "'device_ids' must be an array"}), 400

    for did in device_ids:
        if not _validate_device_id(did):
            return jsonify({"success": False, "error": "Invalid device_id in group"}), 400

    name = name.strip()

    device_ids_set = set(device_ids)

    with _groups_lock:
        if name not in _device_groups and len(_device_groups) >= _MAX_GROUPS:
            return jsonify({"success": False, "error": "Maximum group limit reached"}), 400
        # Remove these devices from any other groups (O(n) with set lookup)
        for gname in list(_device_groups.keys()):
            _device_groups[gname] = [d for d in _device_groups[gname] if d not in device_ids_set]
            if not _device_groups[gname]:
                del _device_groups[gname]
        if device_ids:
            _device_groups[name] = device_ids
        groups_copy = dict(_device_groups)

    socketio.emit("groups_updated", groups_copy)
    return jsonify({"success": True, "groups": groups_copy})


@app.route("/api/groups/<name>", methods=["DELETE"])
def api_groups_delete(name):
    """Delete a device group."""
    with _groups_lock:
        if name in _device_groups:
            del _device_groups[name]
        groups_copy = dict(_device_groups)
    socketio.emit("groups_updated", groups_copy)
    return jsonify({"success": True})


# ---------------------------------------------------------------------------
# EQ API (per-device bass/treble settings)
# ---------------------------------------------------------------------------


@app.route("/api/eq", methods=["POST"])
def api_eq():
    """
    Set EQ (bass/treble) for a device.
    Expects JSON: {"device_id": "...", "bass": -1.0 to 1.0, "treble": -1.0 to 1.0}

    Note: These settings are stored and returned with device info. Actual audio
    EQ processing requires additional audio middleware (e.g., Voicemeeter or
    Equalizer APO) — this API provides the control surface for such integration.
    """
    data = request.get_json(silent=True)
    if not data:
        return jsonify({"success": False, "error": "Invalid JSON"}), 400

    device_id = data.get("device_id")
    bass = data.get("bass")
    treble = data.get("treble")

    if not _validate_device_id(device_id):
        return jsonify({"success": False, "error": "Invalid 'device_id'"}), 400

    try:
        bass = max(-1.0, min(1.0, float(bass))) if bass is not None else 0.0
        treble = max(-1.0, min(1.0, float(treble))) if treble is not None else 0.0
    except (TypeError, ValueError):
        return jsonify({"success": False, "error": "bass/treble must be numbers"}), 400

    with _eq_lock:
        _device_eq[device_id] = {"bass": round(bass, 2), "treble": round(treble, 2)}
        eq = _device_eq[device_id]

    socketio.emit("eq_changed", {"device_id": device_id, "eq": eq})
    return jsonify({"success": True, "device_id": device_id, "eq": eq})


@app.route("/api/eq/<device_id>", methods=["GET"])
def api_eq_get(device_id):
    """Get EQ settings for a device."""
    with _eq_lock:
        eq = _device_eq.get(device_id, {"bass": 0.0, "treble": 0.0})
    return jsonify(eq)


# ---------------------------------------------------------------------------
# Spotify API routes
# ---------------------------------------------------------------------------


@app.route("/api/spotify/login")
def api_spotify_login():
    """Initiate Spotify OAuth 2.0 PKCE login flow."""
    client_id = request.args.get("client_id", "").strip() or _SPOTIFY_CLIENT_ID
    if not client_id:
        return jsonify({"success": False, "error": "No Spotify client_id provided. Set SPOTIFY_CLIENT_ID env var or pass ?client_id=..."}), 400

    verifier, challenge = _spotify_generate_pkce()
    # Store PKCE verifier and client_id in server-side session
    session["spotify_pkce_verifier"] = verifier
    session["spotify_client_id"] = client_id

    params = urllib.parse.urlencode({
        "response_type": "code",
        "client_id": client_id,
        "scope": _SPOTIFY_SCOPES,
        "redirect_uri": _SPOTIFY_REDIRECT_URI,
        "code_challenge_method": "S256",
        "code_challenge": challenge,
    })
    return redirect(f"https://accounts.spotify.com/authorize?{params}")


@app.route("/api/spotify/callback")
def api_spotify_callback():
    """Handle Spotify OAuth callback — exchange code for tokens."""
    code = request.args.get("code")
    error = request.args.get("error")

    if error:
        return f"<h2>Spotify auth error: {html.escape(error)}</h2><p><a href='/'>Back to mixer</a></p>", 400

    if not code:
        return "<h2>Missing authorization code</h2><p><a href='/'>Back to mixer</a></p>", 400

    verifier = session.get("spotify_pkce_verifier")
    client_id = session.get("spotify_client_id", _SPOTIFY_CLIENT_ID)

    if not verifier:
        return "<h2>Session expired — please login again</h2><p><a href='/'>Back to mixer</a></p>", 400

    try:
        resp = http_requests.post("https://accounts.spotify.com/api/token", data={
            "grant_type": "authorization_code",
            "code": code,
            "redirect_uri": _SPOTIFY_REDIRECT_URI,
            "client_id": client_id,
            "code_verifier": verifier,
        }, timeout=10)

        if resp.status_code != 200:
            log.warning("Spotify token exchange failed: %s", resp.text)
            return f"<h2>Token exchange failed</h2><pre>{html.escape(resp.text)}</pre><p><a href='/'>Back to mixer</a></p>", 400

        data = resp.json()
        with _spotify_lock:
            _spotify_tokens["access_token"] = data["access_token"]
            _spotify_tokens["refresh_token"] = data.get("refresh_token")
            _spotify_tokens["expires_at"] = time.time() + data.get("expires_in", 3600) - 60
            _spotify_tokens["client_id"] = client_id

        log.info("Spotify authenticated successfully")
        # Redirect back to the mixer UI
        return redirect("/")

    except Exception as exc:
        log.error("Spotify callback error: %s", exc)
        return f"<h2>Error: {html.escape(str(exc))}</h2><p><a href='/'>Back to mixer</a></p>", 500


@app.route("/api/spotify/status")
def api_spotify_status():
    """Return Spotify connection status."""
    connected = _spotify_token_valid()
    with _spotify_lock:
        np = dict(_spotify_now_playing) if connected else {}
    return jsonify({"connected": connected, "now_playing": np})


@app.route("/api/spotify/now-playing")
def api_spotify_now_playing():
    """Return current Spotify now-playing info."""
    with _spotify_lock:
        np_copy = dict(_spotify_now_playing)
    return jsonify(np_copy)


def _spotify_playback_action(method, endpoint_path):
    """Shared handler for Spotify playback control endpoints."""
    headers = _spotify_get_headers()
    if not headers:
        return jsonify({"success": False, "error": "Not connected to Spotify"}), 401
    try:
        url = f"https://api.spotify.com/v1/me/player/{endpoint_path}"
        resp = getattr(http_requests, method)(url, headers=headers, timeout=5)
        return jsonify({"success": resp.status_code in (200, 204)})
    except Exception as exc:
        return jsonify({"success": False, "error": str(exc)}), 500


@app.route("/api/spotify/play", methods=["POST"])
def api_spotify_play():
    """Resume Spotify playback."""
    return _spotify_playback_action("put", "play")


@app.route("/api/spotify/pause", methods=["POST"])
def api_spotify_pause():
    """Pause Spotify playback."""
    return _spotify_playback_action("put", "pause")


@app.route("/api/spotify/next", methods=["POST"])
def api_spotify_next():
    """Skip to next track."""
    return _spotify_playback_action("post", "next")


@app.route("/api/spotify/previous", methods=["POST"])
def api_spotify_previous():
    """Skip to previous track."""
    return _spotify_playback_action("post", "previous")


@app.route("/api/spotify/disconnect", methods=["POST"])
def api_spotify_disconnect():
    """Disconnect from Spotify (clear tokens)."""
    with _spotify_lock:
        _spotify_tokens["access_token"] = None
        _spotify_tokens["refresh_token"] = None
        _spotify_tokens["expires_at"] = 0
    return jsonify({"success": True})


# ---------------------------------------------------------------------------
# Audio Levels API
# ---------------------------------------------------------------------------


@app.route("/api/levels")
def api_levels():
    """Return current peak audio levels for all devices."""
    with _levels_lock:
        return jsonify(_audio_levels)


# ---------------------------------------------------------------------------
# WebSocket events
# ---------------------------------------------------------------------------


@socketio.on("connect")
def ws_connect():
    """Client connected — send current state."""
    devices = get_playback_devices()
    emit("devices_updated", devices)
    with _groups_lock:
        emit("groups_updated", dict(_device_groups))


@socketio.on("request_refresh")
def ws_request_refresh():
    """Client requested a device refresh via WebSocket."""
    devices = get_playback_devices()
    emit("devices_updated", devices, broadcast=True)


@socketio.on("set_volume")
def ws_set_volume(data):
    """Set volume on a device via WebSocket."""
    device_id = data.get("device_id")
    volume = data.get("volume")
    if _validate_device_id(device_id) and volume is not None:
        try:
            volume = float(volume)
            set_device_volume(device_id, volume)
        except (TypeError, ValueError):
            pass


@socketio.on("set_volumes_batch")
def ws_set_volumes_batch(data):
    """Set batch volumes via WebSocket (lower latency than HTTP)."""
    volumes = data.get("volumes", [])
    if isinstance(volumes, list) and len(volumes) <= 20:
        set_volumes_batch(volumes)


# ---------------------------------------------------------------------------
# Background device monitor (pushes changes via WebSocket)
# ---------------------------------------------------------------------------

_monitor_interval = 3.0  # seconds


def _device_snapshot_key(devices):
    """Return a lightweight hashable key for change detection."""
    return tuple(
        (d["id"], d["name"], d["volume"], d["muted"], d["group"])
        for d in devices
    )


def _device_monitor():
    """Periodically check for device changes and push updates via WebSocket."""
    prev_key = None
    while True:
        try:
            devices = get_playback_devices()
            key = _device_snapshot_key(devices)
            if key != prev_key:
                prev_key = key
                socketio.emit("devices_updated", devices)
        except Exception as exc:
            log.error("Device monitor error: %s", exc)
        socketio.sleep(_monitor_interval)


def _audio_level_monitor():
    """Push per-device audio levels via WebSocket at ~10fps."""
    while True:
        try:
            levels = get_audio_levels()
            if levels:
                socketio.emit("audio_levels", levels)
        except Exception as exc:
            log.error("Audio level monitor error: %s", exc)
        socketio.sleep(0.1)  # ~10fps


def _spotify_monitor():
    """Poll Spotify now-playing and push updates via WebSocket."""
    prev_track = None
    while True:
        try:
            if _spotify_token_valid():
                _spotify_poll_now_playing()
                with _spotify_lock:
                    np = dict(_spotify_now_playing)
                # Only emit when something changes
                track_key = (np.get("track"), np.get("artist"), np.get("is_playing"))
                if track_key != prev_track:
                    prev_track = track_key
                    socketio.emit("spotify_now_playing", np)
        except Exception as exc:
            log.error("Spotify monitor error: %s", exc)
        socketio.sleep(3.0)


# ---------------------------------------------------------------------------
# Application entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    HOST = "127.0.0.1"
    PORT = 5123
    url = f"http://{HOST}:{PORT}"

    log.info("Starting Bluetooth Crossfade Mixer server at %s", url)

    # Start background monitors
    socketio.start_background_task(_device_monitor)
    socketio.start_background_task(_audio_level_monitor)
    socketio.start_background_task(_spotify_monitor)

    # Open the browser after a short delay so the server is ready.
    threading.Timer(1.5, lambda: webbrowser.open(url)).start()

    socketio.run(app, host=HOST, port=PORT, debug=False, allow_unsafe_werkzeug=True)
