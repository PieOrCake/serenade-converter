# Analyse & Auto-Fix

The **Analyse Song** tool (**Tools → Analyse Song**) inspects your arrangement for GW2 playback issues and offers a one-click **Auto-fix** that applies intelligent corrections.

## Running an Analysis

1. Open **Tools → Analyse Song** (or use the menu shortcut)
2. If no melody track is set, you'll be prompted to choose one — see [Melody Track Selection](#melody-track-selection) below
3. A report dialog shows all detected issues per track
4. Click **Auto-fix All Issues** to apply the suggested fixes, or **Close** to dismiss

## Melody Track Selection

When you run Analyse without a melody track set, a dialog appears listing all tracks with two options per track:

- **Melody** (radio button) — select which track carries the main melody. Only one track can be the melody.
- **Preserve** (checkbox) — mark tracks that should be protected from destructive fixes (debass, simplify, density thinning). Preserved tracks keep their notes intact while still receiving octave shifts, clamping, and timing fixes.

The melody track cannot also be preserved (it's already protected by default).

## What the Analysis Detects

For each track, the report flags:

| Issue | Description |
|---|---|
| **Out of range** | Notes outside the selected GW2 instrument's playable range |
| **Octave switches** | Number of octave changes required, and how many are too fast for GW2 to register reliably |
| **Bass duplicates** | Non-melody notes that duplicate the melody pitch class at a lower octave |
| **Dense chords** | Groups of 3+ simultaneous notes that may sound muddy or cause note drops |
| **Tight timing** | Notes too close together for GW2 to play reliably |

Tracks with no issues show a ✓.

## Auto-Fix Phases

Clicking **Auto-fix All Issues** applies corrections in a carefully ordered sequence:

### Phase 1: Octave Shifts (Relative Positioning)

The melody track is shifted first, prioritizing upward shifts to fit the instrument range. A cross-octave cost score ensures the chosen shift doesn't unnecessarily separate melody from accompaniment.

Non-melody tracks are then shifted to **preserve their original pitch offset** from the melody. The arrangement relationship matters more than minimizing out-of-range notes — accompaniment stays below melody even if some notes go out of range, maintaining the musical intent.

### Phase 1.5: Harmonic Substitution

Melody notes that end up just above the instrument range (within one octave) are replaced with a **harmonic substitute**: the note is dropped one octave and a perfect fifth is added. This implies the brightness of the original pitch while staying in range. The substitution only applies when:

- Both the dropped note and fifth fit in range
- Both land in the same GW2 octave (no extra octave switch needed)
- The resulting chord doesn't exceed 4 simultaneous notes

### Phase 1.75: Octave Consolidation

In fast melody passages that rapidly alternate between two GW2 octaves, minority notes are **snapped into the dominant octave** of the run. This eliminates unnecessary octave switching that GW2 may not process fast enough, at the cost of slightly altering the pitch of a few notes.

A run is detected when 4+ melody notes occur within a 1500ms window with 2+ octave crossings.

### Phase 1.9: Cross-Octave Accompaniment Cleanup

GW2 can only play notes in one octave at a time. If an accompaniment note is in a different GW2 octave than the melody at the same moment, the player would have to switch octaves — which can confuse it and cause it to continue playing subsequent melody notes in the wrong octave.

This phase checks every non-melody note that is simultaneous with a melody note. If they're in different GW2 octaves:

1. **Shift** the accompaniment note ±1–2 octaves to land in the melody's octave (if the result is still in instrument range)
2. **Simplify** the note if no valid shift exists (non-preserved tracks only) — simplified notes are hidden but can be restored with Ctrl+Shift+Click

Preserved tracks are shifted when possible but never simplified.

### Phase 2: Track Fixes

Applied per-track:

- **Clamp** (non-melody) — remaining out-of-range notes are moved to the nearest valid pitch, preserving pitch class
- **Debass** (non-melody, non-preserved) — enables bass duplicate removal on the track
- **Simplify** (non-melody, non-preserved) — enables chord simplification (treble + bass)

### Phase 2.5: Density-Adaptive Thinning

Dynamically adjusts the maximum chord size at each beat based on **local melody density**:

| Melody Density | Max Chord Size | Effect |
|---|---|---|
| Sparse (≤2 notes/sec) | 4 | Full chords allowed |
| Moderate (3–5 notes/sec) | 3 | Slightly thinner |
| Dense (6+ notes/sec or octave crossings) | 1 | Melody only |

Octave crossings within the melody count toward density since each costs ~60ms of overhead. Notes are prioritized: **melody > preserved > accompaniment**, and within each tier, higher-pitched notes are kept first.

### Phase 3: Timing Fixes

Applied to **all tracks** including melody. Notes that are too close together for GW2 to play reliably have their durations trimmed to create the minimum required gap. The gap accounts for octave switch overhead when consecutive notes are in different GW2 octaves.

## After Auto-Fix

Auto-fix is a strong starting point, but it won't always produce a perfect arrangement. Complex songs with wide pitch ranges, rapid passages, or intricate harmonies may still need manual adjustment. After running auto-fix:

- **Preview** the result with Play to hear how it sounds
- **Edit individual notes** — move, delete, or redraw notes that don't sound right
- **Adjust tracks** — toggle visibility, change octave assignments, or tweak simplification overrides (Ctrl+Shift+Click)
- **Re-run Analyse** if you make significant changes — it will re-evaluate the current state

Auto-fix is fully undoable with **Ctrl+Z**, so feel free to experiment.

## Preserve Track

Tracks marked as **Preserve** (♫ prefix in the track list) are protected from:

- Bass duplicate removal (debass)
- Chord simplification
- Density-adaptive thinning (muting)

They still receive:

- Octave shifts (relative positioning)
- Clamping of out-of-range notes
- Timing fixes

This is useful for accompaniment tracks with important counter-melodies or harmonies that you don't want simplified away. You can toggle Preserve from the track context menu (right-click a track) or from the Analyse dialog.

## Tips

- **Set a melody track** before running Analyse — it significantly improves auto-fix quality
- **Mark counter-melodies as Preserve** to keep them intact during thinning
- Use **Ctrl+Z** to undo auto-fix if the result isn't what you expected
- Run Analyse multiple times — each run re-evaluates the current state
- After auto-fix, use **Play** to preview the result before exporting
- The log panel shows a detailed summary of every fix applied
