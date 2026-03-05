"""
Bluetooth Speaker Crossfade Mixer — Backend
Flask API + pycaw device enumeration + WASAPI loopback audio routing.
"""

import sys
import os
import webbrowser
import threading
import time
import queue

from flask import Flask, render_template, jsonify, request

# pycaw / COM imports for Windows Core Audio
import comtypes
from pycaw.pycaw import AudioUtilities

# Audio routing imports
import numpy as np
import pyaudiowpatch as pyaudio


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
    """Check if a pycaw AudioDevice is a Bluetooth A2DP (stereo) device.

    Uses PKEY_Device_EnumeratorName ({A45C254E-...}, 24) which is
    'BTHENUM' for A2DP Bluetooth and 'BTHHFENUM' for Hands-Free.
    We only include A2DP — Hands-Free is low-quality mono meant for
    calls and duplicates the same physical speaker.
    """
    if not hasattr(device, 'properties') or not device.properties:
        return False
    # Check the enumerator name property directly
    ENUMERATOR_KEY = "{A45C254E-DF1C-4EFD-8020-67D146A850E0} 24"
    enumerator = device.properties.get(ENUMERATOR_KEY, "")
    if isinstance(enumerator, str):
        upper = enumerator.upper()
        if upper == "BTHHFENUM":
            return False
        if upper == "BTHENUM":
            return True
    # Fallback: scan properties but exclude Hands-Free
    has_bt = False
    for value in device.properties.values():
        if isinstance(value, str):
            v = value.upper()
            if "BTHHFENUM" in v or "HANDS-FREE" in v:
                return False
            if "BTHENUM" in v or "BTH" in v:
                has_bt = True
    return has_bt


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
            print(f"[enumerate] Skipping device: {exc}")
            continue

    return devices_info


def set_device_volume(device_id, volume):
    """
    Set the master volume on the device identified by *device_id*.
    Returns True on success, False on failure.
    """
    _init_com()
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
# Audio Router — WASAPI loopback capture → multi-device output
# ---------------------------------------------------------------------------


class AudioRouter:
    """Captures system audio via WASAPI loopback and streams to BT speakers."""

    CHUNK = 1024        # frames per buffer (~21ms at 48000Hz)
    FORMAT = pyaudio.paFloat32
    NUMPY_DTYPE = np.float32

    def __init__(self):
        self._pa = None
        self._running = False
        self._capture_thread = None
        self._output_threads = {}       # pycaw_device_id -> thread
        self._output_streams = {}       # pycaw_device_id -> pyaudio stream
        self._volumes = {}              # pycaw_device_id -> float 0.0-1.0
        self._lock = threading.Lock()
        self._start_lock = threading.Lock()  # prevent concurrent start/stop
        self._audio_queues = {}         # pycaw_device_id -> queue.Queue
        self._sample_rate = 48000
        self._channels = 2
        self._loopback_info = None
        self._device_index_map = {}     # pycaw_device_id -> pyaudio device index

    def start(self, bt_devices):
        """Start audio routing to the given Bluetooth devices.

        Parameters
        ----------
        bt_devices : list[dict]
            From get_bluetooth_speakers(), each with 'id' and 'name'.
        """
        if not self._start_lock.acquire(blocking=False):
            return False  # Another start/stop in progress
        try:
            return self._start_impl(bt_devices)
        finally:
            self._start_lock.release()

    def _start_impl(self, bt_devices):
        if self._running:
            self._stop_impl()

        try:
            self._pa = pyaudio.PyAudio()
        except Exception as exc:
            print(f"[AudioRouter] Failed to initialize PyAudio: {exc}")
            return False

        # Find a working WASAPI loopback device
        self._loopback_info = self._find_loopback_device()
        if not self._loopback_info:
            print("[AudioRouter] No working WASAPI loopback device found")
            self._cleanup_pa()
            return False

        self._sample_rate = int(self._loopback_info["defaultSampleRate"])
        self._channels = self._loopback_info["maxInputChannels"]

        # Match pycaw BT devices to PyAudio output device indices
        self._device_index_map = self._match_devices(bt_devices)

        if not self._device_index_map:
            print("[AudioRouter] No BT devices matched in PyAudio device list")
            self._cleanup_pa()
            return False

        # Initialize queues and preserve existing volumes
        for dev_id in self._device_index_map:
            self._volumes.setdefault(dev_id, 1.0)
            self._audio_queues[dev_id] = queue.Queue(maxsize=50)

        self._running = True

        # Start output threads first (they block waiting on queues)
        for dev_id, pa_index in self._device_index_map.items():
            t = threading.Thread(
                target=self._output_worker,
                args=(dev_id, pa_index),
                daemon=True,
            )
            self._output_threads[dev_id] = t
            t.start()

        # Start capture thread
        self._capture_thread = threading.Thread(
            target=self._capture_worker,
            daemon=True,
        )
        self._capture_thread.start()

        matched_names = [f"  - {d['name']}" for d in bt_devices if d['id'] in self._device_index_map]
        print(f"[AudioRouter] Started: capturing '{self._loopback_info['name']}' "
              f"({self._sample_rate}Hz {self._channels}ch)")
        for name in matched_names:
            print(f"[AudioRouter]   -> {name}")
        return True

    def stop(self):
        """Stop all audio threads and clean up."""
        with self._start_lock:
            self._stop_impl()

    def _stop_impl(self):
        self._running = False

        # Unblock output threads
        for q in self._audio_queues.values():
            try:
                q.put(None, block=False)
            except queue.Full:
                pass

        if self._capture_thread and self._capture_thread.is_alive():
            self._capture_thread.join(timeout=2)

        for t in self._output_threads.values():
            if t.is_alive():
                t.join(timeout=2)

        for stream in self._output_streams.values():
            try:
                stream.stop_stream()
                stream.close()
            except Exception:
                pass

        self._output_streams.clear()
        self._output_threads.clear()
        self._audio_queues.clear()
        self._device_index_map.clear()
        self._cleanup_pa()
        print("[AudioRouter] Stopped")

    def set_volume(self, device_id, volume):
        """Update the volume multiplier for a device's output stream."""
        with self._lock:
            self._volumes[device_id] = max(0.0, min(1.0, float(volume)))

    def update_devices(self, bt_devices):
        """Re-sync with current BT device list. Restart if devices changed."""
        if not self._pa or not self._running:
            self.start(bt_devices)
            return

        # Don't rebuild the device map if we can't acquire the lock
        if not self._start_lock.acquire(blocking=False):
            return
        try:
            old_ids = set(self._device_index_map.keys())
            new_map = self._match_devices(bt_devices)
            new_ids = set(new_map.keys())

            if old_ids != new_ids:
                print("[AudioRouter] Device change detected, restarting...")
                self._stop_impl()
        finally:
            self._start_lock.release()

        if old_ids != new_ids:
            self.start(bt_devices)

    @property
    def is_running(self):
        return self._running

    @property
    def active_outputs(self):
        return len(self._device_index_map)

    def _find_loopback_device(self):
        """Find the first valid WASAPI loopback device."""
        if not self._pa:
            return None

        for i in range(self._pa.get_device_count()):
            try:
                info = self._pa.get_device_info_by_index(i)
                if (info.get("isLoopbackDevice", False)
                        and info.get("maxInputChannels", 0) > 0
                        and info.get("defaultSampleRate", 0) > 0
                        and info.get("name", "")):
                    return info
            except Exception:
                continue
        return None

    def _match_devices(self, bt_devices):
        """Match pycaw BT devices to PyAudio output device indices.

        BT devices appear in MME (not WASAPI) with truncated endpoint IDs
        as their names. We match by comparing pycaw device IDs to MME names.
        """
        result = {}
        if not self._pa:
            return result

        # Find the MME host API index
        mme_index = None
        for i in range(self._pa.get_host_api_count()):
            info = self._pa.get_host_api_info_by_index(i)
            if info.get("name", "") == "MME":
                mme_index = i
                break

        if mme_index is None:
            return result

        # Collect all MME output devices
        mme_outputs = []
        for i in range(self._pa.get_device_count()):
            try:
                info = self._pa.get_device_info_by_index(i)
                if (info.get("hostApi") == mme_index
                        and info.get("maxOutputChannels", 0) > 0
                        and info.get("name", "")):
                    mme_outputs.append(info)
            except Exception:
                continue

        for bt_dev in bt_devices:
            bt_id = bt_dev["id"]        # e.g. "{0.0.0.00000000}.{7ad893b6-69eb-...}"
            bt_name = bt_dev["name"].lower()

            for mme_dev in mme_outputs:
                mme_name = mme_dev["name"]

                # Match 1: MME name is a truncated prefix of the pycaw endpoint ID
                # MME names are max 31 chars, so "{0.0.0.00000000}.{7ad893b6-69eb"
                # matches pycaw ID "{0.0.0.00000000}.{7ad893b6-69eb-...}"
                if mme_name.startswith("{") and bt_id.lower().startswith(mme_name.lower()):
                    result[bt_id] = mme_dev["index"]
                    break

                # Match 2: friendly name match (case-insensitive substring)
                mme_lower = mme_name.lower()
                if bt_name in mme_lower or mme_lower in bt_name:
                    result[bt_id] = mme_dev["index"]
                    break

        return result

    def _capture_worker(self):
        """Thread: capture loopback audio and distribute to output queues."""
        try:
            stream = self._pa.open(
                format=self.FORMAT,
                channels=self._channels,
                rate=self._sample_rate,
                input=True,
                input_device_index=self._loopback_info["index"],
                frames_per_buffer=self.CHUNK,
            )
        except Exception as exc:
            print(f"[AudioRouter] Failed to open loopback stream: {exc}")
            self._running = False
            return

        print("[AudioRouter] Capture thread running")

        try:
            while self._running:
                try:
                    data = stream.read(self.CHUNK, exception_on_overflow=False)
                except Exception as exc:
                    if self._running:
                        print(f"[AudioRouter] Capture read error: {exc}")
                    time.sleep(0.01)
                    continue

                # Distribute to all output queues
                for dev_id, q in list(self._audio_queues.items()):
                    try:
                        q.put_nowait(data)
                    except queue.Full:
                        # Drop oldest to prevent lag buildup
                        try:
                            q.get_nowait()
                        except queue.Empty:
                            pass
                        try:
                            q.put_nowait(data)
                        except queue.Full:
                            pass
        finally:
            try:
                stream.stop_stream()
                stream.close()
            except Exception:
                pass
            print("[AudioRouter] Capture thread stopped")

    def _output_worker(self, device_id, pa_device_index):
        """Thread: read from queue, apply volume, write to output device."""
        try:
            pa_info = self._pa.get_device_info_by_index(pa_device_index)
            out_channels = min(self._channels, pa_info.get("maxOutputChannels", 2))

            stream = self._pa.open(
                format=self.FORMAT,
                channels=out_channels,
                rate=self._sample_rate,
                output=True,
                output_device_index=pa_device_index,
                frames_per_buffer=self.CHUNK,
            )
            self._output_streams[device_id] = stream
        except Exception as exc:
            print(f"[AudioRouter] Failed to open output {pa_device_index}: {exc}")
            return

        dev_name = pa_info.get("name", str(pa_device_index))
        print(f"[AudioRouter] Output thread running for '{dev_name}'")

        q = self._audio_queues.get(device_id)
        if not q:
            return

        try:
            while self._running:
                try:
                    data = q.get(timeout=0.5)
                except queue.Empty:
                    continue

                if data is None:
                    break

                with self._lock:
                    vol = self._volumes.get(device_id, 1.0)

                if vol < 0.001:
                    # Muted — write silence to keep stream alive
                    try:
                        stream.write(b'\x00' * len(data))
                    except Exception:
                        pass
                    continue

                # Convert to numpy, apply volume, clip
                audio = np.frombuffer(data, dtype=self.NUMPY_DTYPE).copy()

                # Handle channel mismatch
                if self._channels != out_channels and out_channels > 0:
                    audio = audio.reshape(-1, self._channels)[:, :out_channels].flatten()

                audio *= vol
                np.clip(audio, -1.0, 1.0, out=audio)

                try:
                    stream.write(audio.tobytes())
                except Exception as exc:
                    if self._running:
                        print(f"[AudioRouter] Output write error: {exc}")
                    time.sleep(0.01)
        finally:
            try:
                stream.stop_stream()
                stream.close()
            except Exception:
                pass
            print(f"[AudioRouter] Output thread stopped for '{dev_name}'")

    def _cleanup_pa(self):
        """Terminate the PyAudio instance."""
        if self._pa:
            try:
                self._pa.terminate()
            except Exception:
                pass
            self._pa = None


# Global audio router instance
audio_router = AudioRouter()

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
    # Also update the audio router's stream volume
    audio_router.set_volume(device_id, volume)

    if ok:
        return jsonify({"success": True})
    else:
        return jsonify({"success": False, "error": "Failed to set volume"}), 500


@app.route("/api/refresh", methods=["POST"])
def api_refresh():
    """Re-scan devices and return the updated list. Also sync the router."""
    devices = get_bluetooth_speakers()
    if devices and not audio_router.is_running:
        threading.Thread(
            target=lambda: audio_router.start(devices),
            daemon=True,
        ).start()
    elif devices and audio_router.is_running:
        threading.Thread(
            target=lambda: audio_router.update_devices(devices),
            daemon=True,
        ).start()
    return jsonify(devices)


@app.route("/api/router/status", methods=["GET"])
def api_router_status():
    """Return the current state of the audio router."""
    return jsonify({
        "running": audio_router.is_running,
        "outputs": audio_router.active_outputs,
    })


# ---------------------------------------------------------------------------
# Application entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    HOST = "127.0.0.1"
    PORT = 5123
    url = f"http://{HOST}:{PORT}"

    print(f"Starting Bluetooth Crossfade Mixer server at {url}")
    print("Press Ctrl+C to stop.\n")

    # Auto-start audio routing in a background thread with retry
    def start_router():
        _init_com()
        for attempt in range(3):
            devices = get_bluetooth_speakers()
            if devices:
                time.sleep(0.5)  # Let pycaw COM calls settle
                ok = audio_router.start(devices)
                if ok:
                    return
                print(f"[AudioRouter] Start attempt {attempt + 1} failed, retrying...")
                time.sleep(2)
            else:
                print("[AudioRouter] No BT speakers found, retrying in 5s...")
                time.sleep(5)
        print("[AudioRouter] Could not start after 3 attempts. "
              "Connect BT speakers and click Refresh.")

    threading.Thread(target=start_router, daemon=True).start()

    # Open the browser after a short delay so the server is ready.
    threading.Timer(1.5, lambda: webbrowser.open(url)).start()

    app.run(host=HOST, port=PORT, debug=False)
