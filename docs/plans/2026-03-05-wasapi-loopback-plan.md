# WASAPI Loopback Audio Routing — Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Build WASAPI loopback audio routing directly into the app so audio plays through all Bluetooth speakers simultaneously without Voicemeeter or Stereo Mix.

**Architecture:** A capture thread reads system audio via PyAudioWPatch WASAPI loopback from the default output device. Per-device output threads write the captured audio to each Bluetooth speaker with volume scaling applied from the crossfade mixer. The AudioRouter class manages all threads and integrates with the existing Flask backend.

**Tech Stack:** PyAudioWPatch (PortAudio fork with WASAPI loopback), numpy (audio buffer math), threading

---

### Task 1: Add dependencies

**Files:**
- Modify: `requirements.txt`

**Step 1: Update requirements.txt**

```
flask>=3.0
pycaw>=20240210
comtypes>=1.4
pyinstaller>=6.0
PyAudioWPatch>=0.2.12
numpy>=1.26
```

**Step 2: Install dependencies**

Run: `pip install PyAudioWPatch numpy`
Expected: Successfully installed

**Step 3: Verify PyAudioWPatch can see devices**

Run: `python -m pyaudiowpatch`
Expected: List of audio devices including WASAPI loopback entries

**Step 4: Commit**

```bash
git add requirements.txt
git commit -m "feat: add PyAudioWPatch and numpy dependencies for WASAPI loopback"
```

---

### Task 2: Implement AudioRouter class in server.py

**Files:**
- Modify: `server.py`

This is the core task. The `AudioRouter` class manages:
- A loopback capture thread that reads system audio from the default output device
- Per-device output threads that write captured audio to each BT speaker
- Volume scaling per output stream based on crossfade mixer state
- Device matching between pycaw device IDs and PyAudioWPatch device indices

**Step 1: Add imports to server.py**

Add after the existing imports at the top of `server.py`:

```python
import numpy as np
import pyaudiowpatch as pyaudio
import queue
import time
```

**Step 2: Implement the AudioRouter class**

Add before the Flask routes section in `server.py`. The class needs:

```python
class AudioRouter:
    """Captures system audio via WASAPI loopback and streams it to multiple output devices."""

    CHUNK = 1024        # frames per buffer (~23ms at 44100Hz)
    FORMAT = pyaudio.paFloat32
    NUMPY_DTYPE = np.float32

    def __init__(self):
        self._pa = None                 # PyAudio instance
        self._running = False
        self._capture_thread = None
        self._output_threads = {}       # device_id -> thread
        self._output_streams = {}       # device_id -> pyaudio stream
        self._volumes = {}              # device_id -> float (0.0-1.0)
        self._lock = threading.Lock()
        self._audio_queues = {}         # device_id -> queue.Queue
        self._sample_rate = 44100
        self._channels = 2
        self._loopback_device = None
        self._device_index_map = {}     # pycaw device_id -> pyaudio device index

    def start(self, bt_devices):
        """Start the audio router with given Bluetooth device list from pycaw.

        Parameters
        ----------
        bt_devices : list[dict]
            List from get_bluetooth_speakers(), each with 'id' and 'name' keys.
        """
        if self._running:
            self.stop()

        try:
            self._pa = pyaudio.PyAudio()
        except Exception as exc:
            print(f"[AudioRouter] Failed to initialize PyAudio: {exc}")
            return False

        # Find the default WASAPI loopback device
        try:
            self._loopback_device = self._pa.get_default_wasapi_loopback()
        except Exception as exc:
            print(f"[AudioRouter] No WASAPI loopback device found: {exc}")
            self._cleanup_pa()
            return False

        self._sample_rate = int(self._loopback_device["defaultSampleRate"])
        self._channels = self._loopback_device["maxInputChannels"]

        # Build mapping from pycaw device names to PyAudioWPatch output device indices
        self._build_device_map(bt_devices)

        if not self._device_index_map:
            print("[AudioRouter] No BT devices matched in PyAudio device list")
            self._cleanup_pa()
            return False

        # Initialize volumes to 1.0 and create queues
        for dev_id in self._device_index_map:
            self._volumes[dev_id] = self._volumes.get(dev_id, 1.0)
            self._audio_queues[dev_id] = queue.Queue(maxsize=50)

        self._running = True

        # Start output threads first (they block on queue.get)
        for dev_id, pa_index in self._device_index_map.items():
            t = threading.Thread(
                target=self._output_worker,
                args=(dev_id, pa_index),
                daemon=True,
                name=f"output-{dev_id[:20]}"
            )
            self._output_threads[dev_id] = t
            t.start()

        # Start capture thread
        self._capture_thread = threading.Thread(
            target=self._capture_worker,
            daemon=True,
            name="loopback-capture"
        )
        self._capture_thread.start()

        print(f"[AudioRouter] Started: capturing from '{self._loopback_device['name']}' "
              f"-> {len(self._device_index_map)} output(s)")
        return True

    def stop(self):
        """Stop all audio threads and clean up."""
        self._running = False

        # Unblock output threads waiting on queues
        for q in self._audio_queues.values():
            try:
                q.put(None, block=False)
            except queue.Full:
                pass

        # Wait for threads to finish
        if self._capture_thread and self._capture_thread.is_alive():
            self._capture_thread.join(timeout=2)

        for t in self._output_threads.values():
            if t.is_alive():
                t.join(timeout=2)

        # Close output streams
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
        """Re-sync with current Bluetooth device list. Restart if devices changed."""
        if not self._pa:
            self.start(bt_devices)
            return

        # Build new map and compare
        old_ids = set(self._device_index_map.keys())
        new_map = self._match_devices(bt_devices)
        new_ids = set(new_map.keys())

        if old_ids != new_ids:
            print(f"[AudioRouter] Device change detected, restarting...")
            self.stop()
            self.start(bt_devices)

    @property
    def is_running(self):
        return self._running

    @property
    def active_outputs(self):
        return len(self._device_index_map)

    def _build_device_map(self, bt_devices):
        """Match pycaw Bluetooth devices to PyAudioWPatch output device indices."""
        self._device_index_map = self._match_devices(bt_devices)

    def _match_devices(self, bt_devices):
        """Return dict mapping pycaw device_id -> PyAudioWPatch device index.

        Matches by comparing device names (case-insensitive substring match).
        Only includes WASAPI output devices (hostApi matches WASAPI).
        """
        result = {}
        if not self._pa:
            return result

        # Find the WASAPI host API index
        wasapi_index = None
        for i in range(self._pa.get_host_api_count()):
            info = self._pa.get_host_api_info_by_index(i)
            if "WASAPI" in info.get("name", ""):
                wasapi_index = i
                break

        if wasapi_index is None:
            return result

        # Get all PyAudio devices
        pa_devices = []
        for i in range(self._pa.get_device_count()):
            try:
                info = self._pa.get_device_info_by_index(i)
                pa_devices.append(info)
            except Exception:
                continue

        for bt_dev in bt_devices:
            bt_name = bt_dev["name"].lower()
            bt_id = bt_dev["id"]

            for pa_dev in pa_devices:
                # Must be WASAPI and an output device (not a loopback virtual input)
                if pa_dev.get("hostApi") != wasapi_index:
                    continue
                if pa_dev.get("maxOutputChannels", 0) < 1:
                    continue
                # Check if the loopback flag is set (skip loopback devices)
                if pa_dev.get("isLoopbackDevice", False):
                    continue

                pa_name = pa_dev.get("name", "").lower()

                # Substring match: BT device name in PyAudio device name or vice versa
                # e.g. pycaw: "Headphones (i-box Dawn Stereo)"
                #      PyAudio: "Headphones (i-box Dawn Stereo)"
                if bt_name in pa_name or pa_name in bt_name:
                    result[bt_id] = pa_dev["index"]
                    break
                # Also try matching just the core name part
                # Extract text in parentheses for comparison
                for name in [bt_name, pa_name]:
                    start = name.find("(")
                    end = name.find(")")
                    if start != -1 and end != -1:
                        core = name[start+1:end].strip()
                        other = pa_name if name == bt_name else bt_name
                        if core and core in other:
                            result[bt_id] = pa_dev["index"]
                            break
                if bt_id in result:
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
                input_device_index=self._loopback_device["index"],
                frames_per_buffer=self.CHUNK,
            )
        except Exception as exc:
            print(f"[AudioRouter] Failed to open loopback stream: {exc}")
            self._running = False
            return

        print(f"[AudioRouter] Capture started: {self._sample_rate}Hz, {self._channels}ch")

        try:
            while self._running:
                try:
                    data = stream.read(self.CHUNK, exception_on_overflow=False)
                except Exception as exc:
                    print(f"[AudioRouter] Capture read error: {exc}")
                    time.sleep(0.01)
                    continue

                # Distribute to all output queues
                for dev_id, q in list(self._audio_queues.items()):
                    try:
                        q.put_nowait(data)
                    except queue.Full:
                        # Drop oldest frame to prevent growing lag
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
            print("[AudioRouter] Capture stopped")

    def _output_worker(self, device_id, pa_device_index):
        """Thread: read from queue, apply volume, write to output device."""
        try:
            pa_info = self._pa.get_device_info_by_index(pa_device_index)
            # Use the output device's preferred sample rate if different
            out_rate = int(pa_info.get("defaultSampleRate", self._sample_rate))
            out_channels = min(self._channels, pa_info.get("maxOutputChannels", 2))

            stream = self._pa.open(
                format=self.FORMAT,
                channels=out_channels,
                rate=self._sample_rate,  # Use capture rate for consistency
                output=True,
                output_device_index=pa_device_index,
                frames_per_buffer=self.CHUNK,
            )
            self._output_streams[device_id] = stream
        except Exception as exc:
            print(f"[AudioRouter] Failed to open output for device {pa_device_index}: {exc}")
            return

        print(f"[AudioRouter] Output started for device index {pa_device_index}")

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
                    break  # Shutdown signal

                # Apply volume scaling
                with self._lock:
                    vol = self._volumes.get(device_id, 1.0)

                if vol < 0.001:
                    # Muted — write silence to keep stream alive
                    silence = b'\x00' * len(data)
                    try:
                        stream.write(silence)
                    except Exception:
                        pass
                    continue

                # Convert to numpy, scale, convert back
                audio = np.frombuffer(data, dtype=self.NUMPY_DTYPE).copy()

                # Handle channel mismatch: if capture has more channels than output
                if self._channels != out_channels and out_channels > 0:
                    # Reshape to (frames, channels) and take first out_channels
                    audio = audio.reshape(-1, self._channels)[:, :out_channels].flatten()

                audio *= vol
                np.clip(audio, -1.0, 1.0, out=audio)

                try:
                    stream.write(audio.tobytes())
                except Exception as exc:
                    print(f"[AudioRouter] Output write error on {device_id[:20]}: {exc}")
                    time.sleep(0.01)
        finally:
            try:
                stream.stop_stream()
                stream.close()
            except Exception:
                pass
            print(f"[AudioRouter] Output stopped for device index {pa_device_index}")

    def _cleanup_pa(self):
        """Terminate the PyAudio instance."""
        if self._pa:
            try:
                self._pa.terminate()
            except Exception:
                pass
            self._pa = None
```

**Step 3: Create global router instance and integrate with Flask**

Add after the `AudioRouter` class definition, before the Flask routes:

```python
# Global audio router instance
audio_router = AudioRouter()
```

**Step 4: Modify the `/api/volume` endpoint**

Update `api_volume()` to also set the audio router's volume:

After the line `ok = set_device_volume(device_id, volume)`, add:
```python
    audio_router.set_volume(device_id, volume)
```

**Step 5: Add `/api/router/status` endpoint**

```python
@app.route("/api/router/status", methods=["GET"])
def api_router_status():
    """Return the current state of the audio router."""
    return jsonify({
        "running": audio_router.is_running,
        "outputs": audio_router.active_outputs,
    })
```

**Step 6: Modify the `__main__` block to start the router**

In the `if __name__ == "__main__":` block, after the `print` statements and before `app.run(...)`, add:

```python
    # Auto-start audio routing
    def start_router():
        _init_com()
        devices = get_bluetooth_speakers()
        if devices:
            audio_router.start(devices)
        else:
            print("[AudioRouter] No BT speakers found on startup, will retry when devices appear")

    threading.Timer(2.0, start_router).start()
```

**Step 7: Update the device poll to sync the router**

Modify `api_devices()` to update the router when devices change:

```python
@app.route("/api/devices", methods=["GET"])
def api_devices():
    """Return the current list of active playback devices as JSON."""
    devices = get_bluetooth_speakers()
    # Sync router with current device list
    if devices and not audio_router.is_running:
        threading.Thread(target=lambda: audio_router.start(devices), daemon=True).start()
    elif devices and audio_router.is_running:
        audio_router.update_devices(devices)
    return jsonify(devices)
```

**Step 8: Commit**

```bash
git add server.py
git commit -m "feat: add AudioRouter with WASAPI loopback for multi-speaker output"
```

---

### Task 3: Update frontend — add streaming status, remove setup guide

**Files:**
- Modify: `templates/index.html`

**Step 1: Replace the setup guide section**

Remove the entire `<details class="setup-guide">...</details>` block (lines ~184-218) and replace with a streaming status indicator:

```html
    <div id="router-status" class="router-status">
        <span class="status-dot"></span>
        <span id="router-text">Starting audio router...</span>
    </div>
```

**Step 2: Add CSS for the router status**

Add before the closing `</style>` tag:

```css
        /* ── Router status indicator ────────────────────────── */
        .router-status {
            margin-top: 1.5rem;
            display: flex;
            align-items: center;
            gap: 0.6rem;
            font-size: 0.85rem;
            color: #8888aa;
        }

        .router-status .status-dot {
            width: 8px;
            height: 8px;
            border-radius: 50%;
            background: #555;
            transition: background 0.3s;
        }

        .router-status.active .status-dot {
            background: #00ff88;
            box-shadow: 0 0 6px rgba(0, 255, 136, 0.5);
        }

        .router-status.inactive .status-dot {
            background: #ff6b6b;
        }
```

**Step 3: Add JS to poll router status**

Add inside the `<script>` block, after the `fetchDevices` function:

```javascript
    /**
     * Poll the audio router status and update the indicator.
     */
    async function fetchRouterStatus() {
        try {
            const res  = await fetch("/api/router/status");
            const data = await res.json();
            const el   = document.getElementById("router-status");
            const text = document.getElementById("router-text");
            if (data.running) {
                el.className = "router-status active";
                text.textContent = "Streaming to " + data.outputs + " speaker" + (data.outputs !== 1 ? "s" : "");
            } else {
                el.className = "router-status inactive";
                text.textContent = "Audio router inactive — connect speakers";
            }
        } catch (err) {
            // Ignore — status bar already shows connection errors
        }
    }
```

**Step 4: Add router status polling to initialization**

In the initialization section at the bottom of the script, add:

```javascript
    fetchRouterStatus();
    setInterval(fetchRouterStatus, 3000);
```

**Step 5: Commit**

```bash
git add templates/index.html
git commit -m "feat: add streaming status indicator, remove setup guide"
```

---

### Task 4: Update PyInstaller build to bundle PortAudio DLL

**Files:**
- Modify: `build.py`

**Step 1: Update build.py to find and include the PortAudio DLL**

PyAudioWPatch ships with a PortAudio DLL that PyInstaller may not auto-detect. Update `build.py`:

```python
"""
Build script for packaging BluetoothCrossfadeMixer as a standalone .exe
using PyInstaller.

Usage:
    python build.py

The resulting executable will be in the dist/ folder.
"""

import subprocess
import sys
import os

# Find the PyAudioWPatch portaudio DLL path
pyaudio_dll = None
try:
    import pyaudiowpatch
    pyaudio_dir = os.path.dirname(pyaudiowpatch.__file__)
    for f in os.listdir(pyaudio_dir):
        if f.lower().endswith(".dll") and "portaudio" in f.lower():
            pyaudio_dll = os.path.join(pyaudio_dir, f)
            break
except ImportError:
    pass

cmd = [
    sys.executable, "-m", "PyInstaller",
    "--onefile",
    "--console",
    "--name", "BluetoothCrossfadeMixer",
    "--add-data", "templates;templates",
]

if pyaudio_dll:
    cmd.extend(["--add-binary", f"{pyaudio_dll};."])
    print(f"Including PortAudio DLL: {pyaudio_dll}")

cmd.append("server.py")

subprocess.run(cmd, check=True)
```

**Step 2: Rebuild the .exe**

Run: `python build.py`
Expected: Build succeeds, .exe in `dist/` includes the PortAudio DLL

**Step 3: Commit**

```bash
git add build.py
git commit -m "feat: include PortAudio DLL in PyInstaller bundle"
```

---

### Task 5: Test end-to-end and rebuild .exe

**Step 1: Kill any running server**

```bash
# Find and kill processes on port 5123
```

**Step 2: Start the server**

Run: `python server.py`
Expected:
- Server starts at http://127.0.0.1:5123
- AudioRouter auto-starts and shows capture/output messages
- Browser opens with the mixer UI

**Step 3: Verify in the UI**

- Check that the streaming status shows "Streaming to N speakers"
- Drag the control point and verify volume changes on the speakers
- Play music and verify audio comes from all connected BT speakers

**Step 4: Rebuild the .exe**

Run: `python build.py`
Expected: Build completes, .exe in `dist/`

**Step 5: Final commit**

```bash
git add -A
git commit -m "feat: complete WASAPI loopback audio routing integration"
```
