"""
Bluetooth Speaker Crossfade Mixer — Backend
Flask API + pycaw device enumeration for per-device volume control.
"""

import sys
import os
import webbrowser
import threading
from flask import Flask, render_template, jsonify, request

# pycaw / COM imports for Windows Core Audio
import comtypes
from pycaw.pycaw import AudioUtilities


def get_base_dir():
    """Get the base directory - works both in dev and when frozen by PyInstaller."""
    if getattr(sys, 'frozen', False):
        # Running as PyInstaller bundle
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))


BASE_DIR = get_base_dir()

# ---------------------------------------------------------------------------
# Flask app setup
# ---------------------------------------------------------------------------

app = Flask(__name__, template_folder=os.path.join(BASE_DIR, "templates"))

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


def _is_render_device(device):
    """Check if device is a render (playback) device by its endpoint ID prefix."""
    return device.id and device.id.startswith("{0.0.0.")


def _is_bluetooth_device(device):
    """Check if a pycaw AudioDevice is a Bluetooth device.

    Uses PKEY_Device_EnumeratorName ({A45C254E-...}, 24) which is
    'BTHENUM' for A2DP Bluetooth and 'BTHHFENUM' for Hands-Free.
    """
    if not hasattr(device, 'properties') or not device.properties:
        return False
    # Check the enumerator name property directly
    ENUMERATOR_KEY = "{A45C254E-DF1C-4EFD-8020-67D146A850E0} 24"
    enumerator = device.properties.get(ENUMERATOR_KEY, "")
    if isinstance(enumerator, str) and enumerator.upper().startswith("BTH"):
        return True
    # Fallback: scan all string property values for BT indicators
    for value in device.properties.values():
        if isinstance(value, str) and "BTH" in value.upper():
            return True
    return False


def get_bluetooth_speakers():
    """
    Enumerate active Bluetooth audio render (playback) devices and return
    a list of dicts with id, name, and current volume.
    """
    _init_com()
    devices_info = []

    try:
        all_devices = AudioUtilities.GetAllDevices()
    except Exception as exc:
        print(f"[enumerate] Failed to list devices: {exc}")
        return devices_info

    for device in all_devices:
        try:
            # Only include render (output) devices
            if not _is_render_device(device):
                continue

            # Only include Bluetooth devices
            if not _is_bluetooth_device(device):
                continue

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
            print(f"[enumerate] Skipping device: {exc}")
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
        all_devices = AudioUtilities.GetAllDevices()
    except Exception as exc:
        print(f"[set_volume] Failed to enumerate devices: {exc}")
        return False

    for device in all_devices:
        try:
            if device.id == device_id:
                endpoint_volume = device.EndpointVolume
                if endpoint_volume is None:
                    print(f"[set_volume] Device has no volume interface: {device_id}")
                    return False
                endpoint_volume.SetMasterVolumeLevelScalar(volume, None)
                return True
        except Exception as exc:
            print(f"[set_volume] Error setting volume on {device_id}: {exc}")
            return False

    print(f"[set_volume] Device not found: {device_id}")
    return False


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


@app.route("/api/refresh", methods=["POST"])
def api_refresh():
    """Re-scan devices and return the updated list."""
    devices = get_bluetooth_speakers()
    return jsonify(devices)


# ---------------------------------------------------------------------------
# Application entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    HOST = "127.0.0.1"
    PORT = 5123
    url = f"http://{HOST}:{PORT}"

    print(f"Starting Bluetooth Crossfade Mixer server at {url}")
    print("Press Ctrl+C to stop.\n")

    # Open the browser after a short delay so the server is ready.
    threading.Timer(1.5, lambda: webbrowser.open(url)).start()

    app.run(host=HOST, port=PORT, debug=False)
