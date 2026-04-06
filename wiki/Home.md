# Serenade Music Converter

**Serenade Music Converter** is a GUI tool for converting MIDI and MusicXML files into AHK (AutoHotkey) scripts that can be played on instruments in Guild Wars 2 using the [Serenade](https://github.com/PieOrCake/serenade) addon.

![Main Window](screenshots/main_window_context_menu.png)

## What It Does

- Loads MIDI, MusicXML, or existing AHK files
- Displays notes on an interactive **piano roll** where you can edit, add, delete, and rearrange them
- Converts the result to an AHK script that plays the song in GW2 using keyboard inputs
- Supports all GW2 instruments (Piano, Harp, Lute, Horn, Bell, Flute, Bass, etc.)

## Quick Start

1. Download the latest AppImage from [Releases](https://github.com/PieOrCake/serenade-converter/releases)
2. Make it executable: `chmod +x Serenade_Music_Converter-x86_64.AppImage`
3. Run it: `./Serenade_Music_Converter-x86_64.AppImage`
4. **File → Import** (Ctrl+O) to load a MIDI file
5. Edit tracks and notes as needed
6. **File → Save AHK** (Ctrl+S) to export

## User Guide

- **[Getting Started](Getting-Started)** — loading files, basic workflow
- **[Piano Roll Editing](Piano-Roll-Editing)** — selecting, drawing, moving, and resizing notes
- **[Track Management](Track-Management)** — visibility, melody, preserve, octave tools, split/merge
- **[Analyse & Auto-Fix](Analyse-and-Auto-Fix)** — automated analysis and one-click fix for GW2 playback issues
- **[Chord Simplification](Chord-Simplification)** — per-track simplification to reduce complexity
- **[Conversion Settings](Conversion-Settings)** — transpose, chord window, instruments, chord mode
- **[Playback](Playback)** — previewing your arrangement
- **[Keyboard Shortcuts](Keyboard-Shortcuts)** — full shortcut reference
- **[Building from Source](Building-from-Source)** — running and building locally
