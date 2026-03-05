# Bluetooth Speaker Crossfade Mixer — Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Build a Windows app that crossfades audio volume between multiple Bluetooth speakers using a radial spatial mixer UI.

**Architecture:** Python Flask backend uses pycaw to enumerate Bluetooth audio devices and control per-device volume. A single-page web frontend served by Flask provides a radial mixer where speakers are arranged in a circle and a draggable control point determines volume blend via inverse-distance weighting.

**Tech Stack:** Python 3, Flask, pycaw, comtypes, vanilla HTML/CSS/JS with Canvas API.

---

### Task 1: Project Setup

**Files:**
- Create: `requirements.txt`

**Step 1: Create requirements.txt**

```
flask>=3.0
pycaw>=20240210
comtypes>=1.4
```

**Step 2: Install dependencies**

Run: `pip install -r requirements.txt`
Expected: All packages install successfully.

---

### Task 2: Backend — Device Enumeration

**Files:**
- Create: `server.py`

**Step 1: Write the device enumeration module**

Create `server.py` with:
- Imports for pycaw (`AudioUtilities`, `IAudioEndpointVolume`), comtypes, and Flask
- A `get_bluetooth_speakers()` function that:
  - Calls `AudioUtilities.GetAllDevices()` to list all audio endpoints
  - Filters for active render (playback) devices
  - Returns a list of dicts: `{"id": device_id, "name": friendly_name, "volume": current_volume}`
  - Wraps in try/except so individual device failures don't crash the scan
- A `set_device_volume(device_id, volume)` function that:
  - Finds the device by ID from the enumerated endpoints
  - Gets the `IAudioEndpointVolume` interface
  - Calls `SetMasterVolumeLevelScalar(volume, None)` with value clamped 0.0–1.0
  - Returns True on success, False on failure

**Step 2: Verify device enumeration works**

Add a temporary `if __name__ == "__main__"` block that calls `get_bluetooth_speakers()` and prints the result.

Run: `python server.py`
Expected: Prints a list of detected audio playback devices with names and current volumes. If no Bluetooth speakers are connected, an empty list is fine — verify it doesn't crash.

---

### Task 3: Backend — Flask API

**Files:**
- Modify: `server.py`

**Step 1: Add Flask app and routes**

Add to `server.py`:
- `app = Flask(__name__)` with template folder set to `templates`
- `GET /` — serves `index.html` template
- `GET /api/devices` — returns JSON list from `get_bluetooth_speakers()`
- `POST /api/volume` — accepts JSON `{"device_id": "...", "volume": 0.0-1.0}`, calls `set_device_volume()`, returns success/error JSON
- `POST /api/refresh` — re-scans devices and returns updated list
- App launch at bottom: `app.run(host='127.0.0.1', port=5123, debug=False)`
- On startup, print the URL and optionally open the default browser via `webbrowser.open()`

**Step 2: Verify API works**

Run: `python server.py`
Then in another terminal: `curl http://127.0.0.1:5123/api/devices`
Expected: JSON array of devices (may be empty if no speakers connected). Server doesn't crash.

---

### Task 4: Frontend — HTML Shell and Radial Canvas

**Files:**
- Create: `templates/index.html`

**Step 1: Create the HTML page structure**

Create `templates/index.html` with:
- Dark-themed page with centered layout
- Title: "Bluetooth Crossfade Mixer"
- A `<canvas>` element (500x500) for the radial mixer
- A status bar area below showing connected device count
- A "Refresh Devices" button
- A collapsible "Setup Guide" section (collapsed by default) with instructions for:
  - Method 1: Windows "Listen to this device" (Stereo Mix) — step-by-step
  - Method 2: Voicemeeter as alternative
- Basic CSS inline in `<style>`:
  - Dark background (`#1a1a2e`), light text
  - Canvas centered with a subtle border/shadow
  - Button styling consistent with dark theme
  - Collapsible section using `<details>/<summary>`

**Step 2: Verify page loads**

Run: `python server.py`
Open browser to `http://127.0.0.1:5123`
Expected: Dark page with title, empty canvas, refresh button, and collapsible setup guide visible. No JS errors in console.

---

### Task 5: Frontend — Device Fetching and Circle Layout

**Files:**
- Modify: `templates/index.html`

**Step 1: Add JavaScript for device fetching and drawing**

Add `<script>` block with:
- `state` object: `{ devices: [], controlPoint: {x: 250, y: 250}, isDragging: false }`
- `fetchDevices()` — calls `GET /api/devices`, updates `state.devices`, redraws
- `drawMixer()` function that:
  - Clears the canvas
  - Draws a subtle circular boundary (the mixer ring)
  - Positions each device evenly around the circle (angle = `i * 2π / n`)
  - Draws each device as a circle with its name label below
  - Draws the control point as a distinct draggable dot in the center
  - If no devices found, draws "No speakers detected" text in center
- Call `fetchDevices()` on page load
- Set up a `setInterval` to refresh devices every 5 seconds

**Step 2: Verify devices appear on the circle**

Run: `python server.py`, open browser.
Expected: If Bluetooth speakers are connected, they appear as labeled nodes around the circle. If none connected, a "No speakers detected" message shows. The control point dot is visible in the center.

---

### Task 6: Frontend — Draggable Control Point and Volume Calculation

**Files:**
- Modify: `templates/index.html`

**Step 1: Add drag interaction**

Add mouse/touch event handlers:
- `mousedown` / `touchstart` on canvas: if click is near the control point (within ~20px), set `isDragging = true`
- `mousemove` / `touchmove`: if dragging, update `controlPoint` position (clamped within the mixer circle boundary), call `updateVolumes()`, call `drawMixer()`
- `mouseup` / `touchend`: set `isDragging = false`

**Step 2: Add volume calculation**

Add `updateVolumes()` function:
- For each device, compute distance from control point to device position
- Apply inverse-square weighting: `weight_i = 1 / (distance_i^2 + epsilon)` where epsilon prevents division by zero (e.g., `epsilon = 100`)
- Normalize: divide each weight by the max weight so the closest speaker is at 1.0
- Apply a minimum threshold: if a speaker's normalized weight < 0.02, treat as 0
- Send each device's volume to `POST /api/volume` (throttled — only send if value changed by > 0.01 since last send)

**Step 3: Verify dragging changes volumes**

Run the app. Drag the control point toward a speaker.
Expected: The closest speaker's system volume increases, farther speakers decrease. Volume changes are reflected in Windows volume mixer.

---

### Task 7: Frontend — Visual Feedback

**Files:**
- Modify: `templates/index.html`

**Step 1: Add glow and volume display to device nodes**

Update `drawMixer()`:
- Each device node gets a radial gradient glow whose intensity/size scales with its current volume (0 = no glow, 1 = full glow)
- Use a color scheme: e.g., teal/cyan (`#00d4ff`) for active glow on dark background
- Draw volume percentage text inside or next to each device node (e.g., "78%")
- Control point: draw with a slight pulsing animation or distinct color (e.g., white with orange border)
- Connection lines: draw semi-transparent lines from control point to each active speaker (opacity proportional to volume)
- Disconnected/unavailable devices: draw in grey with no glow and "N/A" instead of percentage

**Step 2: Add smooth visual transitions**

- Use `requestAnimationFrame` for the draw loop instead of direct redraws
- Interpolate displayed volume values toward target values (lerp at ~0.15 factor per frame) so glow changes smoothly even when volume jumps

**Step 3: Verify visual feedback**

Run the app, drag control point around.
Expected: Speakers glow brighter when closer to control point, dimmer when farther. Volume percentages update. Connection lines change opacity. Smooth visual transitions.

---

### Task 8: Polish — Error States, Status Bar, and Browser Auto-Open

**Files:**
- Modify: `server.py`
- Modify: `templates/index.html`

**Step 1: Add error handling in frontend**

- If `fetchDevices()` fails (server unreachable), show "Connection lost — retrying..." in the status bar
- If a `POST /api/volume` call fails for a specific device, mark that device as unavailable in state and grey it out
- Status bar shows: device count, connection state, last refresh time

**Step 2: Add browser auto-open on startup**

In `server.py`, add:
```python
import webbrowser
import threading

def open_browser():
    webbrowser.open('http://127.0.0.1:5123')

# Before app.run():
threading.Timer(1.0, open_browser).start()
```

**Step 3: Final verification**

Run: `python server.py`
Expected: Browser opens automatically. If Bluetooth speakers are connected, they appear in the radial mixer. Dragging the control point crossfades between them. Setup guide is accessible. Disconnecting a speaker greys it out.

---

## File Summary

| File | Purpose |
|------|---------|
| `requirements.txt` | Python dependencies (flask, pycaw, comtypes) |
| `server.py` | Flask backend — device enumeration, volume control API, serves frontend |
| `templates/index.html` | Radial mixer UI — canvas, drag interaction, volume calculation, visual feedback, setup guide |

## Running

```bash
pip install -r requirements.txt
python server.py
```

Browser opens to `http://127.0.0.1:5123`. Ensure Bluetooth speakers are connected and multi-output is configured per the setup guide.
