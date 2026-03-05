"""
Bluetooth Speaker Crossfade Mixer — Backend
Flask API + pycaw device enumeration for per-device volume control.
"""

import logging
import webbrowser
import threading
from flask import Flask, render_template, jsonify, request

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = logging.getLogger(__name__)

# pycaw / COM imports for Windows Core Audio
import comtypes
from pycaw.pycaw import AudioUtilities
from pycaw.constants import EDataFlow, DEVICE_STATE

# ---------------------------------------------------------------------------
# Flask app setup
# ---------------------------------------------------------------------------

app = Flask(__name__, template_folder="templates")

# ---------------------------------------------------------------------------
# Device enumeration helpers
# ---------------------------------------------------------------------------


def _init_com():
    """Initialize COM for the current thread. Safe to call multiple times."""
    try:
        comtypes.CoInitialize()
    except OSError:
        # COM already initialized on this thread — that is fine.
        pass


def get_bluetooth_speakers():
    """
    Enumerate all active audio *render* (playback) devices and return a list
    of dicts with id, name, and current volume.

    Includes all playback devices (not just Bluetooth) so the app remains
    useful even without BT speakers connected.
    """
    _init_com()
    devices_info = []

    try:
        # Request only active render (playback) devices from the enumerator.
        all_devices = AudioUtilities.GetAllDevices(
            data_flow=EDataFlow.eRender.value,
            device_state=DEVICE_STATE.ACTIVE.value,
        )
    except Exception as exc:
        log.warning("Failed to list devices: %s", exc)
        return devices_info

    for device in all_devices:
        try:
            # Attempt to read the friendly name; skip unnamed devices.
            friendly_name = device.FriendlyName
            if not friendly_name:
                continue

            # Obtain a unique device id string.
            device_id = device.id
            if not device_id:
                continue

            # Try to get the IAudioEndpointVolume interface to read volume.
            # The pycaw AudioDevice property is named "EndpointVolume".
            volume_level = None
            try:
                endpoint_volume = device.EndpointVolume
                if endpoint_volume is not None:
                    volume_level = endpoint_volume.GetMasterVolumeLevelScalar()
            except Exception:
                # Device may not support volume — that is acceptable.
                pass

            devices_info.append({
                "id": device_id,
                "name": friendly_name,
                "volume": round(volume_level, 4) if volume_level is not None else None,
            })

        except Exception as exc:
            # Individual device failure must not crash the scan.
            log.debug("Skipping device: %s", exc)
            continue

    return devices_info


def set_device_volume(device_id, volume):
    """
    Set the master volume on the device identified by *device_id*.

    Parameters
    ----------
    device_id : str
        The Windows endpoint ID string returned by ``get_bluetooth_speakers()``.
    volume : float
        Desired volume level, clamped to the range 0.0 -- 1.0.

    Returns True on success, False on failure.
    """
    _init_com()

    # Clamp to [0.0, 1.0]
    volume = max(0.0, min(1.0, float(volume)))

    try:
        all_devices = AudioUtilities.GetAllDevices(
            data_flow=EDataFlow.eRender.value,
            device_state=DEVICE_STATE.ACTIVE.value,
        )
    except Exception as exc:
        log.warning("Failed to enumerate devices: %s", exc)
        return False

    for device in all_devices:
        try:
            if device.id != device_id:
                continue
        except Exception:
            continue

        # Found the target device — try to set volume.
        try:
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


# ---------------------------------------------------------------------------
# Flask routes
# ---------------------------------------------------------------------------


@app.route("/")
def index():
    """Serve the frontend single-page application."""
    return render_template("index.html")


@app.route("/api/devices", methods=["GET"])
@app.route("/api/refresh", methods=["POST"])
def api_devices():
    """Return the current list of active playback devices as JSON."""
    devices = get_bluetooth_speakers()
    return jsonify(devices)


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

    if device_id is None:
        return jsonify({"success": False, "error": "Missing 'device_id'"}), 400
    if volume is None:
        return jsonify({"success": False, "error": "Missing 'volume'"}), 400

    try:
        volume = float(volume)
    except (TypeError, ValueError):
        return jsonify({"success": False, "error": "'volume' must be a number"}), 400

    ok = set_device_volume(device_id, volume)
    if ok:
        return jsonify({"success": True})
    else:
        return jsonify({"success": False, "error": "Failed to set volume"}), 500


# ---------------------------------------------------------------------------
# Application entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    HOST = "127.0.0.1"
    PORT = 5123
    url = f"http://{HOST}:{PORT}"

    log.info("Starting Bluetooth Crossfade Mixer server at %s", url)
    log.info("Press Ctrl+C to stop.")

    # Open the browser after a short delay so the server is ready.
    threading.Timer(1.5, lambda: webbrowser.open(url)).start()

    app.run(host=HOST, port=PORT, debug=False)
