# WASAPI Loopback Audio Routing — Design

## Problem

Windows only routes audio to one output device at a time. The crossfade mixer can control per-device volume, but without multi-output routing the audio only plays from whichever speaker is set as the default. Previously this required Stereo Mix setup or Voicemeeter.

## Solution

Build audio routing directly into the app using `sounddevice` (PortAudio) with WASAPI loopback capture. The app captures system audio from the default device and re-streams it to all detected Bluetooth speakers with per-stream volume scaling.

## Architecture

### AudioRouter class (in server.py)

Manages the capture and output threads:

1. **Loopback capture thread** — Uses `sounddevice.InputStream` with WASAPI loopback on the default output device. Reads ~20ms chunks into a shared ring buffer.

2. **Per-device output threads** — One `sounddevice.OutputStream` per Bluetooth speaker. Each reads from the shared buffer, multiplies by the crossfade volume, and writes to its device.

3. **Volume integration** — `/api/volume` updates the router's per-stream gain multiplier. pycaw system volume is no longer used (the router handles volume directly on the audio stream).

### Data flow

```
Default Speaker (system audio)
    |
    v  WASAPI loopback capture
[Ring Buffer] --> Output Thread (BT Speaker A) * volume_A
              --> Output Thread (BT Speaker B) * volume_B
```

### Behavior

- Auto-starts on app launch
- Auto-detects BT speakers and spawns output threads
- When devices change (poll every 5s), spawns/stops threads accordingly
- ~20-40ms buffer for low latency
- Falls back to pycaw volume-only if loopback capture fails

## Files Changed

| File | Change |
|------|--------|
| `requirements.txt` | Add `sounddevice`, `numpy` |
| `server.py` | Add `AudioRouter` class, integrate with startup and volume endpoint |
| `templates/index.html` | Add streaming status indicator, remove setup guide |
| `build.py` | Add portaudio DLL to PyInstaller bundle |

## Dependencies

- `sounddevice` — Python wrapper for PortAudio, supports WASAPI loopback
- `numpy` — Required by sounddevice for audio buffer manipulation
