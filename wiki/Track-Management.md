# Track Management

The **Tracks** panel on the left shows all tracks loaded from the MIDI file. Each track has a checkbox for visibility and a color-coded label.

## Visibility

- **Check/uncheck** a track to show or hide its notes on the piano roll
- Use the **All** / **None** buttons to quickly toggle all tracks
- Hidden tracks are excluded from conversion and playback

## Track Context Menu

Right-click any track to access track tools:

![Track Context Menu](screenshots/main_window_context_menu.png)

| Action | Description |
|---|---|
| **Octave Up / Down** | Shift all notes in the track up or down by one octave (12 semitones) |
| **Clamp to Octave** | Move all notes into a specific octave range (High, Mid, or Low) |
| **Smart Octave Assignment** | Automatically distribute notes across octaves to minimize octave changes |
| **Split by Pitch** | Split a track into two based on pitch (useful for separating treble and bass from a single-track MIDI) |
| **Merge Checked Tracks** | Combine all checked tracks into one |
| **♪ Set as Melody** | Mark this track as the melody track (shown with ♪ prefix) |
| **♫ Preserve Track** | Protect this track from destructive auto-fix passes (debass, simplify, thinning) — see [Analyse & Auto-Fix](Analyse-and-Auto-Fix#preserve-track) |
| **✂ Simplify (treble + bass)** | Enable chord simplification for this track — see [Chord Simplification](Chord-Simplification) |
| **Select All Notes** | Select every note in this track |
| **Delete Track Notes** | Delete all notes in this track |

## Melody Track

Setting a track as the **melody** track (♪) tells the converter to prioritize its notes. Only one track can be the melody at a time.

When a melody track is set, the converter automatically detects notes on other tracks that:
- Play at the **same time** as a melody note (within 5ms)
- Share the same **pitch class** (e.g. both are a C, or both are an F#)
- Are in a **lower octave** than the melody note

These lower-octave duplicates are hidden (marked as simplified) because they muddy the melody without adding harmonic value. Fill notes with *different* pitch classes are preserved, keeping the arrangement rich while letting the melody cut through.

The log shows how many duplicates were hidden when you set a melody track. You can use **Ctrl+Shift+Click** on any hidden note to manually override the decision (same as chord simplification overrides).

## Preserve Track

Tracks marked as **Preserve** (♫ prefix in the track list) are protected from destructive auto-fix operations — bass duplicate removal, chord simplification, and density-adaptive thinning will skip these tracks. They still receive octave shifts, clamping, and timing fixes.

This is useful for accompaniment tracks with important counter-melodies or harmonies that you want to keep intact when running [Analyse & Auto-Fix](Analyse-and-Auto-Fix). Toggle it from the track context menu or from the Analyse dialog when selecting tracks.

## Track Colors

Each track is assigned a distinct color for easy identification on the piano roll. The track colors in the track list match the note colors in the piano roll.
