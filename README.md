# ClipTool

> Voice-activated screen clipper for Windows. Say "computer clip that" and save the last few minutes of your screen + audio instantly.

## Features

- **Voice activation** — Say "computer clip that" to save a clip hands-free
- **GPU-accelerated capture** — Uses DXGI (same as OBS) for zero-lag recording
- **Variable frame rate** — Perfect audio/video sync at any resolution and FPS
- **Dual audio** — Records microphone + desktop audio simultaneously with per-source volume control
- **High resolution** — Supports 720p, 1080p, 1440p, and native monitor resolution
- **Configurable FPS** — 30, 60, or 120fps capture
- **System tray** — Runs quietly in the background, always ready
- **Hotkey** — `Ctrl+Shift+F9` to clip without speaking
- **Persistent settings** — All preferences saved to `config.json`

## Prerequisites

| Requirement | Notes |
|-------------|-------|
| **Windows 10/11** | Required for DXGI capture |
| **Python 3.10+** | [Download Python](https://www.python.org/downloads/) |
| **FFmpeg** | Must be in PATH or installed via WinGet |
| **Microphone** | For voice commands and mic recording |
| **VB-Cable** *(optional)* | For desktop audio capture ([download](https://vb-audio.com/Cable/)) |

## Installation

### 1. Install Python

Download from [python.org](https://www.python.org/downloads/). Make sure to check **"Add Python to PATH"** during installation.

### 2. Install FFmpeg

**Option A — WinGet (recommended):**
```powershell
winget install Gyan.FFmpeg
```

**Option B — Manual:**
1. Download from [ffmpeg.org](https://ffmpeg.org/download.html)
2. Extract to `C:\ffmpeg\`
3. Add `C:\ffmpeg\bin` to your system PATH

### 3. Install VB-Cable *(optional, for desktop audio)*

Download from [vb-audio.com/Cable](https://vb-audio.com/Cable/) and run the installer. This creates a virtual audio device that lets ClipTool record your desktop audio (game sounds, music, etc.).

### 4. Clone and install dependencies

```powershell
git clone https://github.com/Sexigt/clip-tool.git
cd clip-tool
pip install -r requirements.txt
```

### 5. Run

```powershell
python main.py
```

The app starts minimized in the system tray. The main window appears after a few seconds.

## Usage

### Voice Commands

Speak naturally after the wake word **"computer"**:

| Say | Action |
|-----|--------|
| "computer clip that" | Save the last clip (uses default duration) |
| "computer clip last 30 seconds" | Save the last 30 seconds |
| "computer clip last 2 minutes" | Save the last 2 minutes |

The app plays a `clipped.mp3` sound when a clip starts saving.

### Manual Controls

| Button | Action |
|--------|--------|
| **Clip That** | Save using the default duration |
| **1min / 2min / 5min** | Save that many seconds from the buffer |
| **Open Folder** | Open the clips folder in Explorer |

### Hotkey

Press **`Ctrl+Shift+F9`** anywhere to save a clip (uses default duration).

### System Tray

Right-click the tray icon for quick access:
- **Show** — Bring up the main window
- **Clip 1min / 5min** — Quick save
- **Open Folder** — Browse saved clips
- **Quit** — Exit the app

## Settings

### Capture

| Setting | Options | Default | Notes |
|---------|---------|---------|-------|
| **Resolution** | 720p / 1080p / 1440p / Native | 1080p | Native uses your monitor's full resolution |
| **FPS** | 30 / 60 / 120 | 60 | Higher = smoother but more RAM |
| **Quality** | 70 / 85 / 95 | 95 | JPEG quality during capture. Higher = sharper |

### Audio

| Setting | Range | Notes |
|---------|-------|-------|
| **Mic volume** | 0–200% | Volume applied to mic audio in saved clip |
| **Desktop volume** | 0–200% | Volume applied to desktop audio in saved clip |
| **Record Mic** | On/Off | Toggle mic recording |
| **Record Desktop** | On/Off | Toggle desktop audio recording |
| **Default duration** | 30s / 1min / 2min / 5min | Duration used by "computer clip that" |

All settings are saved to `~/clip-tool/clips/config.json` and restored on next launch.

## File Structure

```
clip-tool/
├── main.py              # Main application
├── requirements.txt     # Python dependencies
├── audio/               # Sound effects
│   ├── clipped.mp3      # Played when clipping
│   ├── sound-on.mp3     # Played when toggling recording on
│   └── sound-off.mp3    # Played when toggling recording off
└── clips/               # Saved clips (created on first run)
    ├── clip_*.mp4       # Your saved clips
    ├── config.json      # Saved settings
    ├── _temp/           # Temporary files (auto-cleaned)
    └── cliptool.log     # Error log
```

## Memory Usage

ClipTool keeps a ring buffer of captured frames in RAM. Usage depends on resolution and FPS:

| Resolution | FPS | Buffer | ~RAM |
|------------|-----|--------|------|
| 720p | 30 | 5 min | ~0.7 GB |
| 1080p | 60 | 5 min | ~3.6 GB |
| 1440p | 60 | 5 min | ~6.3 GB |
| Native 2K | 120 | 5 min | ~12 GB |

Lower the FPS or buffer duration if you experience memory pressure.

## Troubleshooting

**"ffmpeg not found"**
- Make sure FFmpeg is installed and in your PATH
- Restart your terminal after installing

**Voice commands not detected**
- Check that your microphone is selected in the Mic dropdown
- Speak clearly after "computer" — the wake word needs to be distinct
- Check `cliptool.log` in the clips folder for errors

**Desktop audio not recording**
- Install [VB-Cable](https://vb-audio.com/Cable/)
- Set VB-Cable as your default playback device, or use it as a monitoring device
- Select the loopback device in the Desktop dropdown

**Clip is blurry**
- Increase the Quality setting (95 recommended)
- Make sure Resolution is set to your monitor's native resolution

**Audio out of sync**
- The app uses variable frame rate encoding for proper sync
- If sync is off, try lowering the FPS setting

## Built With

- [bettercam](https://github.com/CreaticDD/bettercam) — DXGI screen capture
- [faster-whisper](https://github.com/SYSTRAN/faster-whisper) — Speech recognition
- [customtkinter](https://github.com/TomSchimansky/CustomTkinter) — UI framework
- [FFmpeg](https://ffmpeg.org/) — Video/audio encoding
- [pyaudiowpatch](https://github.com/s0d3s/pyaudiowpatch) — Audio capture

## License

MIT
