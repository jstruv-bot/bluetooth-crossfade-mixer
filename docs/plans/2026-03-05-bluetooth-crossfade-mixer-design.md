# Bluetooth Speaker Crossfade Mixer — Design

## Overview

A Windows application that allows users to crossfade audio volume between multiple
Bluetooth speakers using a spatial/radial mixer UI. Python backend controls per-device
volume via Windows Core Audio APIs; a web-based frontend provides the visual mixer.

## Architecture

- **Backend:** Python + Flask + pycaw
  - Enumerates active Bluetooth audio playback devices
  - Exposes REST API: `GET /devices`, `POST /volume` (set per-device volume)
  - Serves the frontend HTML/JS/CSS
  - Polls for device changes on an interval + manual refresh endpoint

- **Frontend:** Single-page HTML/CSS/JS served by Flask
  - Radial mixer: speakers arranged in a circle, draggable control point in center
  - Volume per speaker = inverse-distance-weighted from control point position
  - Real-time visual feedback (glow intensity, percentage labels)

- **Communication:** REST (polling from frontend at ~30fps for smooth interaction)

## Radial Mixer UI

- Speakers displayed as labeled nodes evenly spaced on a circle
- Draggable control point in the center
- Volume formula: `weight_i = 1 / distance_i^2`, normalized so max = 1.0
- Visual feedback: glow/color intensity per speaker, volume percentage text
- Disconnected speakers greyed out and excluded from mix
- "Refresh Devices" button for manual re-scan
- Collapsible "Setup Guide" panel explaining how to enable multi-output on Windows

## Audio Control

- Uses pycaw `IAudioEndpointVolume` to set per-device master volume
- Volume updates throttled to ~30fps
- Values clamped to 0.0–1.0

## Multi-Output Routing Limitation

Windows does not natively route the same audio to multiple outputs simultaneously.
Users must configure this separately via:
1. "Listen to this device" in Windows Sound settings (Stereo Mix method)
2. Third-party tools like Voicemeeter

The app includes a setup guide explaining both methods.

## Error Handling

- Devices that disconnect are greyed out and excluded
- If pycaw can't access a device, it's marked unavailable
- Graceful degradation to single-speaker volume control
- Clear error states in the UI

## File Structure

```
server.py           — Python backend (Flask + pycaw)
templates/
  index.html        — Frontend (HTML + inline CSS/JS)
requirements.txt    — Python dependencies
```

## Dependencies

- flask
- pycaw
- comtypes
