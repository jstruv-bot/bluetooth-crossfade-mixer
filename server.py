"""
Bluetooth Speaker Crossfade Mixer — Backend
Flask API + pycaw device enumeration for per-device volume control.
"""

import logging
import webbrowser
import threading
from flask import Flask, render_template, jsonify, request

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

# Thread-local storage for COM initialization tracking
_com_initialized = threading.local()

# ---------------------------------------------------------------------------
# Device enumeration helpers
# ---------------------------------------------------------------------------

# Maximum length for device_id strings to prevent abuse
_MAX_DEVICE_ID_LEN = 512


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
    of dicts with id, name, and current volume.
    """
    devices_info = []

    try:
        all_devices = _get_active_render_devices()
    except Exception as exc:
        log.error("Failed to list devices: %s", exc)
        return devices_info

    for device in all_devices:
        try:
            friendly_name = device.FriendlyName
            device_id = device.id
            if not friendly_name or not device_id:
                continue

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
            })

        except Exception as exc:
            log.warning("Skipping device: %s", exc)
            continue

    return devices_info


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
                return True
        except Exception as exc:
            log.error("Error setting volume on %s: %s", device_id, exc)
            return False

    log.warning("Device not found: %s", device_id)
    return False


def set_volumes_batch(volumes):
    """
    Set volume on multiple devices in a single enumeration pass.

    Parameters
    ----------
    volumes : list[dict]
        Each dict must have ``device_id`` (str) and ``volume`` (float 0-1).

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
        except Exception as exc:
            log.error("Error setting volume on %s: %s", device.id, exc)
            results[device.id] = False

    # Mark devices that were requested but not found
    for did in lookup:
        if did not in results:
            results[did] = False

    return results


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
    """
    Set volume on a specific device.

    Expects JSON body: {"device_id": "...", "volume": 0.0-1.0}
    """
    data = request.get_json(silent=True)
    if not data:
        return jsonify({"success": False, "error": "Invalid or missing JSON body"}), 400

    device_id = data.get("device_id")
    volume = data.get("volume")

    if not isinstance(device_id, str) or len(device_id) > _MAX_DEVICE_ID_LEN:
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
    """
    Set volume on multiple devices in one call.

    Expects JSON body: {"volumes": [{"device_id": "...", "volume": 0.0-1.0}, ...]}
    """
    data = request.get_json(silent=True)
    if not data or "volumes" not in data:
        return jsonify({"success": False, "error": "Expected JSON with 'volumes' array"}), 400

    volumes = data["volumes"]
    if not isinstance(volumes, list) or len(volumes) > 20:
        return jsonify({"success": False, "error": "'volumes' must be an array (max 20)"}), 400

    # Validate entries
    for entry in volumes:
        did = entry.get("device_id")
        vol = entry.get("volume")
        if not isinstance(did, str) or len(did) > _MAX_DEVICE_ID_LEN:
            return jsonify({"success": False, "error": f"Invalid device_id in batch"}), 400
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
    return jsonify(get_playback_devices())


# ---------------------------------------------------------------------------
# Application entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    HOST = "127.0.0.1"
    PORT = 5123
    url = f"http://{HOST}:{PORT}"

    log.info("Starting Bluetooth Crossfade Mixer server at %s", url)

    # Open the browser after a short delay so the server is ready.
    threading.Timer(1.5, lambda: webbrowser.open(url)).start()

    app.run(host=HOST, port=PORT, debug=False)
