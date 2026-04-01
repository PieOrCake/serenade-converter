# Conversion Settings

The **Conversion Settings** panel controls how notes are mapped to GW2 instrument keypresses when exporting to AHK.

## Instrument

Select the GW2 instrument from the dropdown at the top of the window. Each instrument has a different playable range (shown as red dashed lines on the piano roll):

| Instrument | Range | Octaves |
|---|---|---|
| Piano (C Major) | C3–B5 | 3 |
| Harp | C3–B5 | 3 |
| Lute | C3–B5 | 3 |
| Horn | C3–B5 | 3 |
| Bell | C3–B5 | 3 |
| Bell2 | C3–B5 | 3 |
| Flute | C3–B5 | 3 |
| Bass | C2–B4 | 3 |

Notes outside the instrument's range will be flagged as "Out of range" in the status bar. Use the track context menu's **Clamp to Octave** or **Smart Octave Assignment** to fix these.

## Transpose

Controls pitch shifting before conversion:

- **Auto (maximize white keys)** — automatically finds the transpose value that maps the most notes to white keys (natural notes), reducing the number of sharp/flat keypresses
- **Manual offsets** — shift by a fixed number of semitones (-6 to +6)

## Chord Window

Controls how close in time two notes must be to be considered simultaneous (a chord):

- **Off (simultaneous only)** — only notes starting at exactly the same time are grouped
- **5ms, 10ms, 15ms, 20ms** — notes starting within this window are grouped into a chord

A wider window catches more "near-simultaneous" notes but may group notes that should be sequential.

## Use GW2 Chord Mode

When enabled, the converter detects **major and minor triads** in chord groups and substitutes them with GW2's built-in chord mode keypresses (a single key that plays the full triad). This can reduce the number of keypresses and produce cleaner chords in-game.

GW2's chord mode:
- **Mode 3** — Minor chords
- **Mode 4** — Major chords
- Same key mapping as single notes, but each keypress plays a full triad

## Export

**File → Save AHK** (Ctrl+S) opens the save dialog:

![Save Dialog](screenshots/save__ahk_dialog.png)

The filename is auto-suggested from the Title and Artist fields (e.g., `Piano_Man-Billy_Joel.ahk`). The exported AHK script includes:
- Song metadata (title, artist, instrument)
- Timed `SendInput` commands for each note/chord
- Octave change commands as needed
- Conversion statistics in the header comment
