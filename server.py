"""
Bluetooth Speaker Crossfade Mixer — Backend
Flask API + pycaw device enumeration for per-device volume control.
Features: crossfade curves, mute, presets, groups, WebSocket, EQ, auto-reconnect.
"""

import logging
import webbrowser
import threading
import json
import os
from flask import Flask, render_template, jsonify, request
from flask_socketio import SocketIO, emit

# pycaw / COM imports for Windows Core Audio
import comtypes
from pycaw.pycaw import AudioUtilities
from pycaw.constants import EDataFlow, DEVICE_STATE

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
_MAX_PRESET_NAME_LEN = 64
_MAX_PRESETS = 20
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

            with _muted_lock:
                is_muted = device_id in _muted_devices

            with _eq_lock:
                eq = _device_eq.get(device_id, {"bass": 0.0, "treble": 0.0})

            # Find which group this device belongs to
            group_name = None
            with _groups_lock:
                for gname, members in _device_groups.items():
                    if device_id in members:
                        group_name = gname
                        break

            devices_info.append({
                "id": device_id,
                "name": friendly_name,
                "volume": round(volume_level, 4) if volume_level is not None else None,
                "muted": is_muted,
                "group": group_name,
                "eq": eq,
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

    for device in all_devices:
        try:
            if device.id in lookup:
                endpoint_volume = device.EndpointVolume
                if endpoint_volume is None:
                    results[device.id] = False
                else:
                    endpoint_volume.SetMasterVolumeLevelScalar(lookup[device.id], None)
                    results[device.id] = True
                    # Remember volume for auto-reconnect
                    with _volumes_lock:
                        _last_volumes[device.id] = lookup[device.id]
        except Exception as exc:
            log.error("Error setting volume on %s: %s", device.id, exc)
            results[device.id] = False

    for did in lookup:
        if did not in results:
            results[did] = False

    return results


def _validate_device_id(device_id):
    """Return True if device_id is a valid string within length limits."""
    return isinstance(device_id, str) and 0 < len(device_id) <= _MAX_DEVICE_ID_LEN


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
        return jsonify(_device_groups)


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
            return jsonify({"success": False, "error": f"Invalid device_id in group"}), 400

    name = name.strip()

    with _groups_lock:
        if name not in _device_groups and len(_device_groups) >= _MAX_GROUPS:
            return jsonify({"success": False, "error": "Maximum group limit reached"}), 400
        # Remove these devices from any other groups
        for gname in list(_device_groups.keys()):
            _device_groups[gname] = [d for d in _device_groups[gname] if d not in device_ids]
            if not _device_groups[gname]:
                del _device_groups[gname]
        if device_ids:
            _device_groups[name] = device_ids

    socketio.emit("groups_updated", _device_groups)
    return jsonify({"success": True, "groups": _device_groups})


@app.route("/api/groups/<name>", methods=["DELETE"])
def api_groups_delete(name):
    """Delete a device group."""
    with _groups_lock:
        if name in _device_groups:
            del _device_groups[name]
    socketio.emit("groups_updated", _device_groups)
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


def _device_monitor():
    """Periodically check for device changes and push updates via WebSocket."""
    prev_snapshot = None
    while True:
        try:
            devices = get_playback_devices()
            snapshot = json.dumps(devices, sort_keys=True)
            if snapshot != prev_snapshot:
                prev_snapshot = snapshot
                socketio.emit("devices_updated", devices)
        except Exception as exc:
            log.error("Device monitor error: %s", exc)
        socketio.sleep(_monitor_interval)


# ---------------------------------------------------------------------------
# Application entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    HOST = "127.0.0.1"
    PORT = 5123
    url = f"http://{HOST}:{PORT}"

    log.info("Starting Bluetooth Crossfade Mixer server at %s", url)

    # Start background device monitor
    socketio.start_background_task(_device_monitor)

    # Open the browser after a short delay so the server is ready.
    threading.Timer(1.5, lambda: webbrowser.open(url)).start()

    socketio.run(app, host=HOST, port=PORT, debug=False, allow_unsafe_werkzeug=True)
