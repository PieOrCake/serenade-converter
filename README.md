# Serenade Music Converter

A GUI tool for converting MIDI files to AHK (AutoHotkey) scripts compatible with the [Serenade](https://github.com/PieOrCake/serenade) addon for Guild Wars 2.

![Python](https://img.shields.io/badge/Python-3.10+-blue) ![PyQt6](https://img.shields.io/badge/GUI-PyQt6-green) ![License](https://img.shields.io/badge/License-GPLv3-orange)

## Features

- **Piano Roll Editor** — visual note editing with drag, draw, resize, copy/paste, undo/redo
- **Multi-track support** — load MIDI files with multiple tracks, toggle visibility, set melody track
- **Per-track simplification** — reduce chord complexity by keeping only treble + bass notes (✂)
- **GW2 instrument mapping** — supports all GW2 instruments (Harp, Lute, Horn, Bell, Flute, Bass, etc.)
- **Chord mode** — detects major/minor triads and substitutes GW2 chord keypresses
- **Smart octave detection** — automatically finds the best transpose and octave settings
- **Audio preview** — listen to your arrangement with synthesized playback
- **MusicXML import/export** — interoperate with notation software
- **Batch conversion** — convert multiple MIDI files at once
- **Song submission** — submit your arrangements to the community song index

## Installation

### Using an AppImage Manager (recommended)
1. Install [Gear Lever](https://flathub.org/apps/it.mijorus.gearlever) from Flathub
2. Download the latest `.AppImage` from [Releases](https://github.com/PieOrCake/serenade-converter/releases)
3. Right-click the downloaded file and select Open With → Gear Lever

Gear Lever will integrate the app into your desktop — no terminal needed.

### Manual (terminal)
Download the latest `.AppImage` from [Releases](https://github.com/PieOrCake/serenade-converter/releases), then:

```bash
chmod +x Serenade_Music_Converter-x86_64.AppImage
./Serenade_Music_Converter-x86_64.AppImage
```

### Windows
Download `Serenade.Music.Converter.exe` from [Releases](https://github.com/PieOrCake/serenade-converter/releases) and run it directly. No installation required.

## Running from Source

```bash
# Clone the repo
git clone https://github.com/PieOrCake/serenade-converter.git
cd serenade-converter

# Create a virtual environment and install dependencies
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

# Run
python3 midi2ahk.py
```

### Requirements

- Python 3.10+
- PyQt6
- mido
- numpy
- pygame

## Building the AppImage

Requires [Podman](https://podman.io/) (or Docker with minor script edits).

```bash
./build-appimage.sh
```

This builds inside an Ubuntu 22.04 container using PyInstaller, then packages the result as an AppImage. The output is `Serenade_Music_Converter-x86_64.AppImage`.

## Keyboard Shortcuts

| Shortcut | Action |
|---|---|
| Ctrl+O | Import file |
| Ctrl+S | Save AHK |
| Ctrl+Z / Ctrl+Y | Undo / Redo |
| Ctrl+C / Ctrl+V | Copy / Paste notes |
| Ctrl+A | Select all notes |
| Del | Delete selected notes |
| Ctrl+D | Toggle draw mode |
| Ctrl+= / Ctrl+- | Zoom in / out |
| Ctrl+Shift+Click | Toggle note simplification (on simplified tracks) |
| Mouse wheel | Scroll vertically |
| Shift+Wheel | Scroll horizontally |
| Ctrl+Wheel | Zoom |

## License

This project is licensed under the [GNU General Public License v3.0](LICENSE).
