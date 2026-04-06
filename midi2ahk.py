#!/usr/bin/env python3
"""
midi2ahk — MIDI to GW2 AHK converter for Serenade.

GUI tool (PyQt6) that converts MIDI files to AHK scripts compatible with the
Serenade addon for Guild Wars 2.

GW2 Piano mapping (C Major instruments):
    Key 1-7: C D E F G A B  (natural notes)
    F1-F5:   C# D# F# G# A# (sharps/flats)
    Key 9:   Octave up
    Key 0:   Octave down
    GW2 piano has 3 octaves (low/mid/high).
"""

__version__ = '1.3.0'

import sys
import webbrowser
import os
import re
import json
import urllib.request
import urllib.parse
from bisect import bisect_left, bisect_right
from collections import Counter

# Add venv to path
VENV_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.venv')
VENV_SITE = os.path.join(VENV_DIR, 'lib')
for d in os.listdir(VENV_SITE) if os.path.isdir(VENV_SITE) else []:
    sp = os.path.join(VENV_SITE, d, 'site-packages')
    if os.path.isdir(sp) and sp not in sys.path:
        sys.path.insert(0, sp)

import mido
import xml.etree.ElementTree as ET
import zipfile
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QFileDialog, QComboBox, QLineEdit, QTextEdit,
    QGroupBox, QProgressBar, QMessageBox, QListWidget, QListWidgetItem,
    QSizePolicy, QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView,
    QTabWidget, QCheckBox, QSplitter, QScrollBar, QMenu, QSpinBox,
    QDialog, QDialogButtonBox
)
from PyQt6.QtCore import Qt, QSettings, QTimer, QThread, pyqtSignal, QRectF, QPointF
from PyQt6.QtGui import QFont, QPainter, QColor, QPen, QBrush, QCursor, QPolygonF, QIcon, QKeySequence

# ── GW2 Note Mapping ──────────────────────────────────────────────────────────

NOTE_MAP = {
    0:  ('note', 1),    # C
    1:  ('sharp', 1),   # C# → F1
    2:  ('note', 2),    # D
    3:  ('sharp', 2),   # D# → F2
    4:  ('note', 3),    # E
    5:  ('note', 4),    # F
    6:  ('sharp', 3),   # F# → F3
    7:  ('note', 5),    # G
    8:  ('sharp', 4),   # G# → F4
    9:  ('note', 6),    # A
    10: ('sharp', 5),   # A# → F5
    11: ('note', 7),    # B
}


# Reverse lookup for sorting chord notes by pitch (low to high)
_GW2_KEY_SEMITONE = {}
for _semi, (_kt, _kn) in NOTE_MAP.items():
    _GW2_KEY_SEMITONE[(_kt, _kn)] = _semi
GW2_MID = 1

# Semitones that land on white keys (C major natural notes)
WHITE_KEY_SEMITONES = {0, 2, 4, 5, 7, 9, 11}

SEMITONE_NAMES = ['C', 'C#', 'D', 'D#', 'E', 'F', 'F#', 'G', 'G#', 'A', 'A#', 'B']


def search_itunes(query, limit=5):
    """Search iTunes for song metadata. Returns list of (title, artist, source_label)."""
    results = []
    try:
        encoded = urllib.parse.urlencode({'term': query, 'entity': 'song', 'limit': limit})
        url = f'https://itunes.apple.com/search?{encoded}'
        req = urllib.request.Request(url, headers={'User-Agent': 'Serenade-MIDI-Converter/1.0'})
        with urllib.request.urlopen(req, timeout=5) as resp:
            data = json.loads(resp.read().decode('utf-8'))
        seen = set()
        for item in data.get('results', []):
            title = item.get('trackName', '')
            artist = item.get('artistName', '')
            if title and (title.lower(), artist.lower()) not in seen:
                seen.add((title.lower(), artist.lower()))
                results.append((title, artist, 'iTunes'))
    except Exception:
        pass  # silently fail — suggestions are best-effort
    return results


def generate_metadata_suggestions(filepath, mid=None):
    """Generate multiple (title, artist) suggestions from filename and MIDI metadata.

    Returns a list of (title, artist, source_label) tuples, deduplicated.
    """
    suggestions = []
    seen = set()

    def add(title, artist, source):
        key = (title.lower(), artist.lower())
        if key not in seen:
            seen.add(key)
            suggestions.append((title, artist, source))

    stem = os.path.splitext(os.path.basename(filepath))[0]
    cleaned = stem.replace('_', ' ').strip()
    cleaned = re.sub(r'^\d+[\.\s\-]+\s*', '', cleaned)

    # Online lookup via iTunes Search API
    itunes_results = search_itunes(cleaned)
    for title, artist, source in itunes_results:
        add(title, artist, source)

    # Pattern: "A - B" → suggest both orderings
    if ' - ' in cleaned:
        parts = cleaned.split(' - ', 1)
        a, b = parts[0].strip(), parts[1].strip()
        add(b, a, f'"{a}" as artist')
        add(a, b, f'"{b}" as artist')

    # Pattern: "Title (Artist)" or "Title [Artist]"
    if m := re.match(r'^(.+?)\s*[\(\[]\s*(.+?)\s*[\)\]]$', cleaned):
        add(m.group(1).strip(), m.group(2).strip(), 'parentheses')

    # Pattern: "Title by Artist"
    if m := re.search(r'^(.+?)\s+by\s+(.+)$', cleaned, re.IGNORECASE):
        add(m.group(1).strip(), m.group(2).strip(), '"by" keyword')

    # Raw filename as title, no artist
    add(cleaned, '', 'filename only')

    # MIDI metadata suggestions
    if mid:
        meta = extract_midi_metadata(mid)
        if meta['title']:
            add(meta['title'], meta['artist'], 'MIDI metadata')
        if meta['artist'] and not meta['title']:
            add(cleaned, meta['artist'], 'MIDI artist + filename')
        if meta['copyright']:
            add(cleaned, meta['copyright'], 'MIDI copyright')

    return suggestions


def extract_midi_metadata(mid):
    """Extract metadata from MIDI file (track names, copyright, text events).

    Returns dict with keys: title, artist, copyright, texts.
    """
    meta = {'title': '', 'artist': '', 'copyright': '', 'texts': []}

    # Track/instrument-like names to skip when looking for a real title
    _skip_words = {'piano', 'guitar', 'bass', 'drum', 'drums', 'vocal', 'vocals',
                   'voice', 'melody', 'lead', 'rhythm', 'strings', 'synth', 'organ',
                   'flute', 'harp', 'violin', 'cello', 'brass', 'choir', 'tempo',
                   'rh', 'lh', 'left', 'right'}

    for track in mid.tracks:
        for msg in track:
            if msg.type == 'track_name' and msg.name.strip():
                name = msg.name.strip()
                # Skip names that are just instrument/part labels
                words = set(name.lower().replace('-', ' ').split())
                if not meta['title'] and not words.issubset(_skip_words):
                    meta['title'] = name
            elif msg.type == 'copyright' and hasattr(msg, 'text') and msg.text.strip():
                meta['copyright'] = msg.text.strip()
            elif msg.type == 'text' and hasattr(msg, 'text') and msg.text.strip():
                meta['texts'].append(msg.text.strip())
            elif msg.type == 'instrument_name' and hasattr(msg, 'name'):
                pass  # skip instrument names

    # Try to find artist in text events
    for text in meta['texts']:
        lower = text.lower()
        if any(kw in lower for kw in ['composed by', 'artist:', 'performer:', 'arranged by']):
            # Extract the name after the keyword
            for kw in ['composed by ', 'artist: ', 'performer: ', 'arranged by ']:
                if kw in lower:
                    idx = lower.index(kw) + len(kw)
                    meta['artist'] = text[idx:].strip()
                    break

    return meta


def find_best_transpose(notes_with_times):
    """Find the transposition (0-11 semitones) that minimizes octave boundary
    crossings between consecutive notes.  White-key percentage is used only
    as a tiebreaker when two shifts produce the same number of crossings.

    Returns (best_shift, white_pct) where best_shift is semitones to add.
    """
    if not notes_with_times:
        return 0, 100.0

    midi_notes = [n for _, n, _ in notes_with_times]
    total = len(midi_notes)
    best_shift = 0
    best_crossings = total  # worst case
    best_white = 0

    for shift in range(12):
        crossings = 0
        for i in range(total - 1):
            if (midi_notes[i] + shift) // 12 != (midi_notes[i + 1] + shift) // 12:
                crossings += 1

        white = sum(1 for n in midi_notes if (n + shift) % 12 in WHITE_KEY_SEMITONES)

        # Primary: fewest crossings.  Tiebreaker: most white keys.
        if crossings < best_crossings or (crossings == best_crossings and white > best_white):
            best_crossings = crossings
            best_shift = shift
            best_white = white

    return best_shift, (best_white / total) * 100.0


def midi_note_to_gw2(note, base_octave_midi=48):
    offset = note - base_octave_midi
    if offset < 0 or offset >= 36:
        return None
    octave = offset // 12
    semitone = offset % 12
    # Boundary C notes (semitone 0) can be played as key 8 (C') in the
    # previous octave, avoiding an unnecessary octave switch.
    if semitone == 0 and octave > 0:
        return (octave - 1, 'note', 8)
    key_type, key_num = NOTE_MAP[semitone]
    return (octave, key_type, key_num)


def gw2_key_name(key_type, key_num):
    if key_type == 'note':
        return str(key_num)  # regular number row keys (not numpad)
    else:
        return f'F{key_num}'


# GW2 mode system: 0 increases (octave up), 9 decreases (octave down)
GW2_MODE_LOW = 0
GW2_MODE_MID = 1
GW2_MODE_HIGH = 2
GW2_MODE_MINOR = 3
GW2_MODE_MAJOR = 4


def detect_triad(midi_pitches):
    """Check if a set of MIDI pitches contains a major or minor triad.
    Returns (root_semitone, 'major'|'minor') or None.
    A major triad = root, root+4, root+7 semitones.
    A minor triad = root, root+3, root+7 semitones."""
    if len(midi_pitches) < 3:
        return None
    pcs = set(p % 12 for p in midi_pitches)
    if len(pcs) < 3:
        return None
    # Check each pitch class as potential root
    for root in pcs:
        major_3rd = (root + 4) % 12
        minor_3rd = (root + 3) % 12
        fifth = (root + 7) % 12
        if major_3rd in pcs and fifth in pcs:
            return (root, 'major')
        if minor_3rd in pcs and fifth in pcs:
            return (root, 'minor')
    return None


def get_track_info(midi_file):
    """Return list of (index, name, note_count, note_range, is_duplicate) for each track.

    Also returns the recommended track index (best single track for GW2 piano).
    """
    mid = mido.MidiFile(midi_file)
    tracks = []
    seen_fingerprints = {}  # fingerprint → first track index

    for i, track in enumerate(mid.tracks):
        notes = [msg.note for msg in track if msg.type == 'note_on' and msg.velocity > 0]
        note_count = len(notes)
        name = track.name if track.name else "(unnamed)"
        note_range = (min(notes), max(notes)) if notes else (0, 0)

        # Fingerprint: note count + sorted note tuple (detects exact duplicates)
        fingerprint = (note_count, tuple(sorted(notes))) if notes else None
        is_duplicate = False
        if fingerprint and fingerprint in seen_fingerprints:
            is_duplicate = True
        elif fingerprint:
            seen_fingerprints[fingerprint] = i

        tracks.append((i, name, note_count, note_range, is_duplicate))

    # Recommend best single track: most notes, fits in 3 octaves, not a duplicate
    best_idx = None
    best_score = -1
    for i, name, count, (lo, hi), is_dup in tracks:
        if count == 0 or is_dup:
            continue
        span = hi - lo
        # Score: prefer more notes, penalize range > 36 semitones (3 octaves)
        range_penalty = max(0, span - 36) * 10
        score = count - range_penalty
        if score > best_score:
            best_score = score
            best_idx = i

    return tracks, mid, best_idx


def find_best_base_octave(notes_with_times):
    if not notes_with_times:
        return 48, 0, 0

    midi_notes = [n for _, n, _ in notes_with_times]
    min_note = min(midi_notes)
    max_note = max(midi_notes)

    best_base = 48
    best_count = 0

    start = (max(0, min_note - 35) // 12) * 12
    for base in range(start, min(128, max_note + 1), 12):
        count = sum(1 for n in midi_notes if base <= n < base + 36)
        if count > best_count:
            best_count = count
            best_base = base

    return best_base, best_count, len(midi_notes)


def extract_notes(mid, track_indices=None):
    """Extract note events from MIDI.

    track_indices: list of track indices, or None for all tracks merged.
    """
    if mid.type == 0 or track_indices is None:
        events_raw = []
        abs_time = 0.0
        for msg in mid:
            abs_time += msg.time
            if msg.type == 'note_on' and msg.velocity > 0:
                events_raw.append((abs_time * 1000, msg.note, msg.velocity))
        return events_raw
    else:
        tempo = 500000
        tpb = mid.ticks_per_beat

        tempo_map = []
        if mid.type == 1 and len(mid.tracks) > 0:
            tick = 0
            for msg in mid.tracks[0]:
                tick += msg.time
                if msg.type == 'set_tempo':
                    tempo_map.append((tick, msg.tempo))
        if not tempo_map:
            tempo_map = [(0, tempo)]

        def ticks_to_ms(target_tick):
            ms = 0.0
            current_tick = 0
            current_tempo = tempo_map[0][1]
            tempo_idx = 0
            while current_tick < target_tick:
                next_tempo_tick = tempo_map[tempo_idx + 1][0] if tempo_idx + 1 < len(tempo_map) else target_tick
                end_tick = min(next_tempo_tick, target_tick)
                delta_ticks = end_tick - current_tick
                ms += (delta_ticks / tpb) * (current_tempo / 1000.0)
                current_tick = end_tick
                if current_tick >= next_tempo_tick and tempo_idx + 1 < len(tempo_map):
                    tempo_idx += 1
                    current_tempo = tempo_map[tempo_idx][1]
            return ms

        events = []
        for tidx in track_indices:
            track = mid.tracks[tidx]
            abs_tick = 0
            for msg in track:
                abs_tick += msg.time
                if msg.type == 'note_on' and msg.velocity > 0:
                    events.append((ticks_to_ms(abs_tick), msg.note, msg.velocity))

        events.sort(key=lambda e: e[0])
        return events


# ── MusicXML Parsing ──────────────────────────────────────────────────────────

STEP_TO_SEMI = {'C': 0, 'D': 2, 'E': 4, 'F': 5, 'G': 7, 'A': 9, 'B': 11}


def parse_musicxml_file(filepath):
    """Parse a MusicXML file (.musicxml, .xml, or .mxl compressed).
    Returns the XML root element."""
    import os
    ext = os.path.splitext(filepath)[1].lower()

    if ext == '.mxl':
        with zipfile.ZipFile(filepath, 'r') as z:
            # Try META-INF/container.xml for the rootfile path
            try:
                container = ET.parse(z.open('META-INF/container.xml')).getroot()
                # Handle namespace
                for rf in container.iter():
                    if rf.tag.endswith('rootfile') and rf.get('full-path'):
                        return ET.parse(z.open(rf.get('full-path'))).getroot()
            except (KeyError, ET.ParseError):
                pass
            # Fallback: find any .musicxml or .xml file
            for name in z.namelist():
                if 'META-INF' in name:
                    continue
                if name.endswith('.musicxml') or name.endswith('.xml'):
                    return ET.parse(z.open(name)).getroot()
            raise ValueError("No MusicXML file found in .mxl archive")
    else:
        return ET.parse(filepath).getroot()


def _xml_ns(root):
    """Extract namespace prefix from root element."""
    if root.tag.startswith('{'):
        return root.tag.split('}')[0] + '}'
    return ''


def get_musicxml_parts(filepath):
    """Get part info from a MusicXML file.
    Returns (parts, root, best_idx) matching get_track_info() format."""
    root = parse_musicxml_file(filepath)
    ns = _xml_ns(root)

    # Part names from part-list
    part_names = {}
    part_list = root.find(f'{ns}part-list')
    if part_list is not None:
        for sp in part_list.findall(f'{ns}score-part'):
            pid = sp.get('id')
            pn = sp.find(f'{ns}part-name')
            part_names[pid] = pn.text.strip() if pn is not None and pn.text else '(unnamed)'

    parts = []
    seen_fp = {}

    for idx, part in enumerate(root.findall(f'{ns}part')):
        pid = part.get('id')
        name = part_names.get(pid, f'Part {idx}')

        midi_notes = []
        for note_elem in part.iter(f'{ns}note'):
            if note_elem.find(f'{ns}rest') is not None:
                continue
            pitch = note_elem.find(f'{ns}pitch')
            if pitch is None:
                continue
            step = pitch.find(f'{ns}step').text
            octave = int(pitch.find(f'{ns}octave').text)
            alter_elem = pitch.find(f'{ns}alter')
            alter = int(float(alter_elem.text)) if alter_elem is not None else 0
            midi_notes.append((octave + 1) * 12 + STEP_TO_SEMI[step] + alter)

        note_count = len(midi_notes)
        note_range = (min(midi_notes), max(midi_notes)) if midi_notes else (0, 0)

        fp = (note_count, tuple(sorted(midi_notes))) if midi_notes else None
        is_dup = False
        if fp and fp in seen_fp:
            is_dup = True
        elif fp:
            seen_fp[fp] = idx

        parts.append((idx, name, note_count, note_range, is_dup))

    best_idx = None
    best_score = -1
    for i, name, count, (lo, hi), is_dup in parts:
        if count == 0 or is_dup:
            continue
        span = hi - lo
        score = count - max(0, span - 36) * 10
        if score > best_score:
            best_score = score
            best_idx = i

    return parts, root, best_idx


def extract_notes_musicxml(root, part_indices=None):
    """Extract note events from MusicXML.
    Returns list of (time_ms, midi_note, velocity)."""
    ns = _xml_ns(root)

    all_parts = root.findall(f'{ns}part')
    if part_indices is None:
        part_indices = list(range(len(all_parts)))

    events = []

    for pidx in part_indices:
        part = all_parts[pidx]

        divisions = 1
        tempo_bpm = 120.0
        current_time_ms = 0.0
        prev_note_start = 0.0

        for measure in part.findall(f'{ns}measure'):
            for elem in measure:
                tag = elem.tag.replace(ns, '')

                if tag == 'attributes':
                    div_elem = elem.find(f'{ns}divisions')
                    if div_elem is not None:
                        divisions = int(div_elem.text)

                elif tag == 'direction':
                    sound = elem.find(f'{ns}sound')
                    if sound is not None and sound.get('tempo'):
                        tempo_bpm = float(sound.get('tempo'))

                elif tag == 'note':
                    is_chord = elem.find(f'{ns}chord') is not None
                    is_rest = elem.find(f'{ns}rest') is not None
                    dur_elem = elem.find(f'{ns}duration')
                    duration = int(dur_elem.text) if dur_elem is not None else 0

                    ms_per_div = (60000.0 / tempo_bpm) / divisions

                    if is_chord:
                        note_start = prev_note_start
                    else:
                        note_start = current_time_ms
                        prev_note_start = note_start
                        current_time_ms += duration * ms_per_div

                    if not is_rest:
                        pitch = elem.find(f'{ns}pitch')
                        if pitch is not None:
                            step = pitch.find(f'{ns}step').text
                            octave = int(pitch.find(f'{ns}octave').text)
                            alter_elem = pitch.find(f'{ns}alter')
                            alter = int(float(alter_elem.text)) if alter_elem is not None else 0
                            midi_note = (octave + 1) * 12 + STEP_TO_SEMI[step] + alter
                            events.append((note_start, midi_note, 80))

                elif tag == 'forward':
                    dur_elem = elem.find(f'{ns}duration')
                    if dur_elem is not None:
                        ms_per_div = (60000.0 / tempo_bpm) / divisions
                        current_time_ms += int(dur_elem.text) * ms_per_div

                elif tag == 'backup':
                    dur_elem = elem.find(f'{ns}duration')
                    if dur_elem is not None:
                        ms_per_div = (60000.0 / tempo_bpm) / divisions
                        current_time_ms -= int(dur_elem.text) * ms_per_div

    events.sort(key=lambda e: e[0])
    return events


def extract_musicxml_metadata(root):
    """Extract title and artist from MusicXML."""
    ns = _xml_ns(root)
    meta = {'title': '', 'artist': '', 'copyright': '', 'texts': []}

    work = root.find(f'{ns}work')
    if work is not None:
        wt = work.find(f'{ns}work-title')
        if wt is not None and wt.text:
            meta['title'] = wt.text.strip()

    if not meta['title']:
        mt = root.find(f'{ns}movement-title')
        if mt is not None and mt.text:
            meta['title'] = mt.text.strip()

    ident = root.find(f'{ns}identification')
    if ident is not None:
        for creator in ident.findall(f'{ns}creator'):
            ctype = creator.get('type', '').lower()
            if ctype in ('composer', 'lyricist', 'arranger') and creator.text:
                if not meta['artist']:
                    meta['artist'] = creator.text.strip()
        for rights in ident.findall(f'{ns}rights'):
            if rights.text:
                meta['copyright'] = rights.text.strip()

    return meta


def convert(midi_file, output_file, track_indices=None, title=None, author=None, transpose=None, base_octave=None, chord_window_ms=5, notes_override=None, use_chords=False, instrument=None, smooth_octaves=False):
    """Convert a MIDI file to a GW2 AHK script. Returns (success, log_lines).

    track_indices: list of track indices, or None for all tracks merged.
    transpose: semitones to shift (0-11), or None for auto-detect.
    base_octave: MIDI note for Low octave base, or None for auto-detect.
    chord_window_ms: max gap (ms) between notes to group as a chord.
        Default 5 (nearly simultaneous). Higher values collapse arpeggios.
    use_chords: if True, substitute detected triads with GW2 chord mode keypresses.
    smooth_octaves: if True, flatten short octave excursions in fast passages.
    """
    log = []
    if notes_override is not None:
        notes = notes_override
        log.append(f"Source: MusicXML")
    else:
        try:
            mid = mido.MidiFile(midi_file)
        except Exception as e:
            return False, [f"ERROR: Failed to read MIDI: {e}"]
        log.append(f"Type: {mid.type}, Ticks/beat: {mid.ticks_per_beat}, Tracks: {len(mid.tracks)}")
        notes = extract_notes(mid, track_indices)

    if not notes:
        return False, log + ["ERROR: No notes found in selected track!"]

    # Normalise to 4-tuples: (time_ms, pitch, velocity, is_melody)
    if len(notes[0]) == 3:
        notes = [(t, n, v, False) for t, n, v in notes]

    log.append(f"Note events: {len(notes)}")

    # Auto-transpose to minimize octave boundary crossings
    notes_3 = [(t, n, v) for t, n, v, _m in notes]  # 3-tuple view for helpers
    if transpose is None:
        transpose, white_pct = find_best_transpose(notes_3)
    else:
        _, orig_pct = find_best_transpose(notes_3)
        white_pct = sum(1 for _, n, _, _m in notes if (n + transpose) % 12 in WHITE_KEY_SEMITONES) / len(notes) * 100

    if transpose != 0:
        log.append(f"Transpose: +{transpose} semitones (white keys: {white_pct:.0f}%)")
        notes = [(t, n + transpose, v, m) for t, n, v, m in notes]
    else:
        log.append(f"Transpose: none needed (white keys: {white_pct:.0f}%)")

    if base_octave is not None:
        base_midi = base_octave
        in_range = sum(1 for _, n, _, _m in notes if base_midi <= n < base_midi + 36)
        total = len(notes)
    elif instrument:
        # Use the selected instrument's base octave so GW2 octave mapping
        # matches the range lines shown in the piano roll.
        base_midi = None
        for inst_name, _, inst_low, _ in GW2_INSTRUMENTS:
            if inst_name == instrument:
                base_midi = inst_low
                break
        if base_midi is None:
            base_midi, _, _ = find_best_base_octave(notes_3)
        in_range = sum(1 for _, n, _, _m in notes if base_midi <= n < base_midi + 36)
        total = len(notes)
    else:
        base_midi, in_range, total = find_best_base_octave(notes_3)
    log.append(f"Base octave: MIDI {base_midi} (C{base_midi // 12 - 1})")
    log.append(f"Notes in range: {in_range}/{total}")
    if total > in_range:
        log.append(f"WARNING: {total - in_range} notes out of range (clamped)")

    # Group notes into chords. Each note joins the current group if it's
    # within chord_window_ms of the FIRST note in the group.
    chord_groups = []
    current_group = [notes[0]]
    for i in range(1, len(notes)):
        if notes[i][0] - current_group[0][0] <= chord_window_ms:
            current_group.append(notes[i])
        else:
            chord_groups.append(current_group)
            current_group = [notes[i]]
    chord_groups.append(current_group)
    has_melody = any(m for _, _, _, m in notes)

    log.append(f"Note groups (chords+singles): {len(chord_groups)}")

    # Generate AHK script
    if not title:
        title = os.path.splitext(os.path.basename(midi_file))[0]
    lines = []
    lines.append(f'; title: {title}')
    if author:
        lines.append(f'; author: {author}')
    lines.append(f'; instrument: {instrument or "Piano"}')
    lines.append(f'; Converted by Serenade Music Converter v{__version__}')
    lines.append('')

    # First pass: compute target octave for each group
    # Each gw2_note is (octave, key_type, key_num, is_melody)
    group_data = []
    for group in chord_groups:
        gw2_notes = []
        for _, midi_note, _, is_mel in group:
            result = midi_note_to_gw2(midi_note, base_midi)
            if result is None:
                continue  # drop out-of-range notes instead of clamping
            gw2_notes.append((*result, is_mel))

        if not gw2_notes:
            group_data.append((group, gw2_notes, GW2_MID))
            continue

        # If any melody notes exist in this group, prefer their octave
        melody_in_group = [n for n in gw2_notes if n[3]]
        if has_melody and melody_in_group:
            octave_counts = {}
            for oct, _, _, _ in melody_in_group:
                octave_counts[oct] = octave_counts.get(oct, 0) + 1
            target_octave = max(octave_counts, key=octave_counts.get)
        else:
            octave_counts = {}
            for oct, _, _, _ in gw2_notes:
                octave_counts[oct] = octave_counts.get(oct, 0) + 1
            target_octave = max(octave_counts, key=octave_counts.get)
        group_data.append((group, gw2_notes, target_octave))

    # Second pass (optional): smooth octave sequence using hybrid time + run-length.
    # In fast passages (short gaps), only switch octave if we stay in the
    # new octave for MIN_RUN groups. In slow passages (long gaps), allow
    # octave changes freely since there's enough time for the switch.
    MIN_RUN = 4       # min consecutive groups in fast passages
    SLOW_GAP_MS = 250  # gaps >= this are "slow" — always allow changes
    octave_changes_saved = 0
    if smooth_octaves:
        # Start smoothing from the first group's actual octave
        first_oct = GW2_MID
        for _, notes_check, oct_check in group_data:
            if notes_check:
                first_oct = oct_check
                break
        current_smooth_oct = first_oct
        def _group_has_melody(gd):
            """Return True if a group_data entry contains any melody notes."""
            return any(is_mel for _, _, _, is_mel in gd[1])

        i = 0
        while i < len(group_data):
            grp, notes, oct = group_data[i]
            if not notes:
                i += 1
                continue
            if oct == current_smooth_oct:
                i += 1
                continue

            # Never smooth groups that contain melody notes — the melody
            # must always play at its correct octave.
            if has_melody and _group_has_melody(group_data[i]):
                current_smooth_oct = oct
                i += 1
                continue

            # Check timing gap before this group
            gap_ms = 0
            if i > 0:
                gap_ms = grp[0][0] - group_data[i - 1][0][0][0]

            # Slow passage — allow octave change regardless of run length
            if gap_ms >= SLOW_GAP_MS:
                current_smooth_oct = oct
                i += 1
                continue

            # Fast passage — require minimum run length
            run_length = 0
            for j in range(i, len(group_data)):
                if group_data[j][1]:  # has notes
                    if group_data[j][2] == oct:
                        run_length += 1
                    else:
                        break
            if run_length >= MIN_RUN:
                current_smooth_oct = oct
                i += 1
            else:
                # Flatten this short run to current octave, but skip melody groups
                for j in range(i, len(group_data)):
                    if group_data[j][1]:
                        if group_data[j][2] == oct:
                            if has_melody and _group_has_melody(group_data[j]):
                                break  # don't flatten melody groups
                            group_data[j] = (group_data[j][0], group_data[j][1], current_smooth_oct)
                            octave_changes_saved += 1
                        else:
                            break
                i += 1

        if octave_changes_saved > 0:
            log.append(f"Octave smoothing: flattened {octave_changes_saved} short octave excursions")

    # Count actual octave changes
    octave_change_count = 0
    prev_oct = GW2_MID
    for _, notes, oct in group_data:
        if notes and oct != prev_oct:
            octave_change_count += abs(oct - prev_oct)
            prev_oct = oct
    log.append(f"Octave changes: {octave_change_count}")

    # current_mode: 0=low, 1=mid, 2=high, 3=minor chords, 4=major chords
    current_mode = GW2_MODE_MID
    last_time_ms = 0.0
    chord_subs = 0

    OCTAVE_SWITCH_MS = 60  # per-step delay for GW2 to process octave change

    for group, gw2_notes, target_octave in group_data:
        group_time = group[0][0]
        gap_ms = group_time - last_time_ms

        # Compute how many octave steps are needed for this group
        oct_steps = 0
        if gw2_notes:
            triad = None
            if use_chords and len(gw2_notes) >= 3:
                midi_pitches = [n for _, n, _, _m in group]
                triad = detect_triad(midi_pitches)
            if triad:
                _rs, _ct = triad
                needed = GW2_MODE_MAJOR if _ct == 'major' else GW2_MODE_MINOR
            else:
                needed = target_octave
            oct_steps = abs(needed - current_mode)

        # Emit sleep with octave cost absorbed
        oct_cost = oct_steps * OCTAVE_SWITCH_MS
        rest_ms = gap_ms - oct_cost
        if rest_ms > 1:
            lines.append(f'Sleep, {int(round(rest_ms))}')

        if not gw2_notes:
            last_time_ms = group_time
            continue

        # Check for triad substitution
        triad = None
        if use_chords and len(gw2_notes) >= 3:
            midi_pitches = [n for _, n, _, _m in group]
            triad = detect_triad(midi_pitches)

        if triad:
            root_semi, chord_type = triad
            target_mode = GW2_MODE_MAJOR if chord_type == 'major' else GW2_MODE_MINOR
            # Switch to chord mode
            if target_mode != current_mode:
                if target_mode > current_mode:
                    for _ in range(target_mode - current_mode):
                        lines.append('SendInput {0}')
                        lines.append(f'Sleep, {OCTAVE_SWITCH_MS}')
                else:
                    for _ in range(current_mode - target_mode):
                        lines.append('SendInput {9}')
                        lines.append(f'Sleep, {OCTAVE_SWITCH_MS}')
                current_mode = target_mode
            # Send root key
            key_type, key_num = NOTE_MAP[root_semi]
            lines.append(f'SendInput {{{gw2_key_name(key_type, key_num)}}}')
            chord_subs += 1
        else:
            # Regular note output — ensure we're in an octave mode (0-2)
            target_mode_oct = target_octave  # 0=low, 1=mid, 2=high
            if target_mode_oct != current_mode:
                if target_mode_oct > current_mode:
                    for _ in range(target_mode_oct - current_mode):
                        lines.append('SendInput {0}')
                        lines.append(f'Sleep, {OCTAVE_SWITCH_MS}')
                else:
                    for _ in range(current_mode - target_mode_oct):
                        lines.append('SendInput {9}')
                        lines.append(f'Sleep, {OCTAVE_SWITCH_MS}')
                current_mode = target_mode_oct

            # Prioritise melody notes, then sort low-to-high
            def _note_sort_key(n):
                oct, key_type, key_num, is_mel = n
                return (0 if is_mel else 1, oct * 12 + _GW2_KEY_SEMITONE.get((key_type, key_num), 0))
            gw2_notes_sorted = sorted(gw2_notes, key=_note_sort_key)
            key_names = []
            for oct, key_type, key_num, _mel in gw2_notes_sorted:
                name = gw2_key_name(key_type, key_num)
                if name not in key_names:
                    key_names.append(name)
            key_names = key_names[:4]

            if key_names:
                for k in key_names:
                    lines.append(f'SendInput {{{k}}}')

        last_time_ms = group_time

    # Trailing sleep so the last note has time to sustain
    lines.append('Sleep, 1000')

    if chord_subs > 0:
        log.append(f"Chord substitutions: {chord_subs} triads replaced with GW2 chord mode")

    with open(output_file, 'w') as f:
        f.write('\r\n'.join(lines) + '\r\n')

    duration_s = last_time_ms / 1000.0
    log.append(f"Duration: {int(duration_s // 60)}:{int(duration_s % 60):02d}")
    log.append(f"Output lines: {len(lines)}")
    log.append(f"Saved: {output_file}")

    return True, log


# ── Piano Roll ────────────────────────────────────────────────────────────────

TRACK_COLORS = [
    (0x4F, 0xC3, 0xF7),  # Blue
    (0x81, 0xC7, 0x84),  # Green
    (0xFF, 0xB7, 0x4D),  # Orange
    (0xE5, 0x73, 0x73),  # Red
    (0xBA, 0x68, 0xC8),  # Purple
    (0xFF, 0xD5, 0x4F),  # Yellow
    (0x4D, 0xD0, 0xE1),  # Cyan
    (0xA1, 0x88, 0x7F),  # Brown
]



# ── Audio Synth Engine ────────────────────────────────────────────────────────

import numpy as np
import io
import wave as wave_mod

SYNTH_RATE = 44100

def _midi_to_freq(note):
    return 440.0 * (2.0 ** ((note - 69) / 12.0))

_synth_cache = {}  # pitch -> base waveform (1 second at vel=127)

def _build_piano_wave(pitch, sample_rate=SYNTH_RATE):
    """Pre-build a 2-second piano waveform for the given MIDI pitch."""
    freq = _midi_to_freq(pitch)
    dur = 2.0
    n = int(sample_rate * dur)
    t = np.linspace(0, dur, n, False).astype(np.float32)

    # Inharmonicity by register
    if pitch < 40:
        B = 0.0004
    elif pitch < 60:
        B = 0.00015
    else:
        B = 0.00005

    # Limit harmonics for speed (8 max — still sounds good)
    n_harm = max(3, min(8, int(8000 / max(1, freq))))
    amps =   [1.0, 0.6, 0.35, 0.18, 0.10, 0.06, 0.04, 0.03]
    decays = [1.0, 1.8, 2.6,  3.4,  4.2,  5.0,  5.8,  6.6]

    tone = np.zeros(n, dtype=np.float32)
    for h in range(n_harm):
        idx = h + 1
        f_h = freq * idx * np.sqrt(1 + B * idx * idx)
        if f_h > sample_rate * 0.45:
            break
        phase = (idx * 0.7) % (2 * np.pi)
        harmonic = amps[h] * np.sin((2 * np.pi * f_h) * t + phase, dtype=np.float32)
        harmonic *= np.exp(-decays[h] * t, dtype=np.float32)
        tone += harmonic

    # Hammer noise
    hd = min(int(0.003 * sample_rate), n)
    if hd > 0:
        tone[:hd] += (np.random.randn(hd).astype(np.float32) * 0.06 *
                       np.linspace(1, 0, hd, dtype=np.float32))

    # Double-decay envelope
    env = (0.35 * np.exp(-3.5 * t, dtype=np.float32) +
           0.65 * np.exp(-0.4 * t, dtype=np.float32))
    att = min(int(0.004 * sample_rate), n)
    if att > 0:
        env[:att] *= np.linspace(0, 1, att, dtype=np.float32)
    env[-min(int(0.06 * sample_rate), n):] *= np.linspace(1, 0,
        min(int(0.06 * sample_rate), n), dtype=np.float32)

    tone *= env * 0.30
    return tone

def synth_note(freq, duration_s, velocity=100, sample_rate=SYNTH_RATE):
    """Synthesize an acoustic grand piano tone (cached per pitch)."""
    pitch = int(round(12 * np.log2(freq / 440.0) + 69))
    pitch = max(0, min(127, pitch))

    if pitch not in _synth_cache:
        _synth_cache[pitch] = _build_piano_wave(pitch, sample_rate)

    base = _synth_cache[pitch]
    n_samples = int(sample_rate * duration_s)
    if n_samples == 0:
        return np.zeros(0, dtype=np.float32)

    vel_norm = velocity / 127.0

    if n_samples <= len(base):
        tone = base[:n_samples].copy() * vel_norm
    else:
        # For very long notes, pad with silence (piano decays by 2s anyway)
        tone = np.zeros(n_samples, dtype=np.float32)
        tone[:len(base)] = base * vel_norm

    # Fade out at end
    rel = min(int(0.08 * sample_rate), n_samples)
    if rel > 0 and n_samples > rel:
        tone[-rel:] *= np.linspace(1, 0, rel, dtype=np.float32)

    return tone

def simulate_gw2_playback(notes):
    """Transform notes to simulate GW2 instrument constraints.
    - Flatten velocity (GW2 has no dynamics)
    - All notes in a chord group play (SendInput fires near-instantly)
    - When a melody track is set, melody notes get priority in chord groups
    Accepts 4-tuples (s, d, p, v) or 5-tuples (s, d, p, v, is_melody).
    Returns a new list of (start_ms, duration_ms, pitch, velocity) 4-tuples."""
    if not notes:
        return notes
    GW2_VEL = 100
    # Normalise to 5-tuples
    if len(notes[0]) == 4:
        notes = [(s, d, p, v, False) for s, d, p, v in notes]

    has_melody = any(m for _, _, _, _, m in notes)
    if not has_melody:
        return [(s, d, p, GW2_VEL) for s, d, p, v, _m in notes]

    # Group simultaneous notes (within 5ms) and prioritise melody
    CHORD_WIN = 5
    result = []
    i = 0
    while i < len(notes):
        group = [notes[i]]
        j = i + 1
        while j < len(notes) and notes[j][0] - group[0][0] <= CHORD_WIN:
            group.append(notes[j])
            j += 1
        # Sort: melody first, then by pitch
        group.sort(key=lambda n: (0 if n[4] else 1, n[2]))
        # Keep up to 4 notes per chord, melody prioritised
        for s, d, p, v, _m in group[:4]:
            result.append((s, d, p, GW2_VEL))
        i = j
    return result


def render_notes_to_audio(notes, sample_rate=SYNTH_RATE):
    """Render a list of (start_ms, duration_ms, pitch, velocity) to a numpy audio array."""
    if not notes:
        return np.zeros(0, dtype=np.int16)
    total_ms = max(n[0] + n[1] for n in notes) + 200
    total_samples = int(total_ms / 1000.0 * sample_rate)
    audio = np.zeros(total_samples, dtype=np.float32)
    for start_ms, dur_ms, pitch, vel in notes:
        freq = _midi_to_freq(pitch)
        dur_s = max(0.05, dur_ms / 1000.0)
        tone = synth_note(freq, dur_s, vel, sample_rate)
        start_sample = int(start_ms / 1000.0 * sample_rate)
        end_sample = start_sample + len(tone)
        if end_sample > len(audio):
            tone = tone[:len(audio) - start_sample]
            end_sample = len(audio)
        if start_sample < len(audio):
            audio[start_sample:end_sample] += tone
    # Normalize and convert to int16
    peak = np.max(np.abs(audio))
    if peak > 0:
        audio = audio / peak * 0.85
    return (audio * 32767).astype(np.int16)


class RenderWorker(QThread):
    """Background thread for rendering audio from notes."""
    finished = pyqtSignal(object)  # emits the int16 audio array

    def __init__(self, notes, gw2_preview=False, parent=None):
        super().__init__(parent)
        self._notes = notes
        self._gw2_preview = gw2_preview

    def run(self):
        notes = self._notes
        if self._gw2_preview:
            notes = simulate_gw2_playback(notes)  # returns 4-tuples
        else:
            # Strip melody flag if present (render expects 4-tuples)
            if notes and len(notes[0]) == 5:
                notes = [(s, d, p, v) for s, d, p, v, _m in notes]
        audio = render_notes_to_audio(notes)
        self.finished.emit(audio)

def audio_to_wav_bytes(audio, sample_rate=SYNTH_RATE):
    """Convert int16 numpy array to WAV bytes."""
    buf = io.BytesIO()
    with wave_mod.open(buf, 'wb') as wf:
        wf.setnchannels(1)
        wf.setsampwidth(2)
        wf.setframerate(sample_rate)
        wf.writeframes(audio.tobytes())
    buf.seek(0)
    return buf

_pygame_inited = False
def _ensure_pygame():
    global _pygame_inited
    if not _pygame_inited:
        import pygame
        pygame.mixer.init(frequency=SYNTH_RATE, size=-16, channels=1, buffer=512)
        _pygame_inited = True

def play_preview_note(pitch, velocity=100, duration_s=0.4):
    """Play a single note preview (non-blocking)."""
    import pygame
    _ensure_pygame()
    freq = _midi_to_freq(pitch)
    tone = synth_note(freq, duration_s, velocity)
    audio = (tone * 32767 / max(0.01, np.max(np.abs(tone)))).astype(np.int16)
    wav = audio_to_wav_bytes(audio)
    sound = pygame.mixer.Sound(wav)
    sound.play()


# GW2 instrument definitions: (name, key, midi_low, midi_high)
# midi_low/high = the full playable range including the octave-repeat top note
GW2_INSTRUMENTS = [
    ("Piano",     "C Major", 48, 84),   # C3-C6, 3 octaves
    ("Harp",      "C Major", 48, 84),   # C3-C6, 3 octaves
    ("Lute",      "C Major", 48, 84),   # C3-C6, 3 octaves
    ("Minstrel",  "C Major", 48, 84),   # C3-C6, 3 octaves
    ("Horn",      "C Major", 48, 84),   # C3-C6, 3 octaves
    ("Bell",      "D Major", 50, 86),   # D3-D6, 3 octaves
    ("Verdarach", "D Major", 50, 86),   # D3-D6, 3 octaves
    ("Flute",     "E Major", 64, 88),   # E4-E6, 2 octaves
    ("Bass",      "C Major", 36, 60),   # C2-C4, 2 octaves
]


# GW2 playback timing constraints (from MusicPlayer.h InstrumentProfile)
GW2_MIN_NOTE_DELAY_MS = 50    # Minimum ms between consecutive notes
GW2_OCTAVE_SWAP_DELAY_MS = 60 # Additional delay when octave changes
# Octave boundaries: each octave spans 12 semitones; notes in different
# floor(pitch/12) groups require an octave swap command.

def _note_octave(pitch, base_pitch=60):
    """Return which GW2 octave a pitch falls in (0=low, 1=mid, 2=high)."""
    # Each GW2 octave is 12 semitones. Base pitch is the middle octave root.
    # For C-major instruments: low=C3(48), mid=C4(60), high=C5(72)
    offset = pitch - (base_pitch - 12)  # offset from low octave start
    return max(0, min(2, offset // 12))

def analyze_gw2_issues(notes):
    """Analyze notes for GW2 playback issues. Sets .warning on each note.
    Returns (overlap_count, too_close_count, octave_tight_count).
    """
    # Clear all warnings
    active = [n for n in notes if not n.deleted]
    for n in active:
        n.warning = ''
    if len(active) < 2:
        return (0, 0, 0)

    # Sort by start time, then pitch
    sorted_notes = sorted(active, key=lambda n: (n.start_ms, n.pitch))

    overlap_count = 0
    too_close_count = 0
    octave_tight_count = 0

    for i in range(len(sorted_notes)):
        note = sorted_notes[i]
        if note.warning:
            continue  # already flagged

        # Check against all subsequent notes for overlaps
        for j in range(i + 1, len(sorted_notes)):
            other = sorted_notes[j]
            if other.start_ms >= note.start_ms + note.duration_ms + GW2_MIN_NOTE_DELAY_MS * 2:
                break  # too far ahead

            # Overlap: two notes playing at the same time
            if other.start_ms < note.start_ms + note.duration_ms and abs(other.start_ms - note.start_ms) < 1:
                # Simultaneous notes (chord) - GW2 fires them as rapid keypresses
                if not note.warning:
                    note.warning = 'chord'
                    overlap_count += 1
                if not other.warning:
                    other.warning = 'chord'
                    overlap_count += 1

        # Check gap to next sequential note
        if i + 1 < len(sorted_notes):
            nxt = sorted_notes[i + 1]
            gap = nxt.start_ms - (note.start_ms + note.duration_ms)
            note_oct = _note_octave(note.pitch)
            next_oct = _note_octave(nxt.pitch)
            needs_octave_change = (note_oct != next_oct)

            if needs_octave_change:
                min_gap = GW2_MIN_NOTE_DELAY_MS + GW2_OCTAVE_SWAP_DELAY_MS
                if gap < min_gap and not nxt.warning:
                    nxt.warning = 'octave_tight'
                    octave_tight_count += 1
            else:
                if gap < GW2_MIN_NOTE_DELAY_MS and gap >= 0 and not nxt.warning:
                    nxt.warning = 'too_close'
                    too_close_count += 1

    return (overlap_count, too_close_count, octave_tight_count)

def fix_gw2_issues(notes):
    """Auto-fix GW2 playback issues. Returns number of fixes applied.
    - Trims note durations to ensure minimum gaps between sequential notes
    - Adds extra gap padding for octave changes
    - Does NOT touch simultaneous notes (chords) — converter handles those natively
    """
    active = [n for n in notes if not n.deleted]
    if len(active) < 2:
        return 0

    # Sort by start time, then pitch
    sorted_notes = sorted(active, key=lambda n: (n.start_ms, n.pitch))

    # Group simultaneous notes (within 1ms = same chord)
    groups = []
    i = 0
    while i < len(sorted_notes):
        group = [sorted_notes[i]]
        j = i + 1
        while j < len(sorted_notes) and abs(sorted_notes[j].start_ms - sorted_notes[i].start_ms) < 1:
            group.append(sorted_notes[j])
            j += 1
        groups.append(group)
        i = j

    fixes = 0

    for gi in range(len(groups) - 1):
        cur_group = groups[gi]
        nxt_group = groups[gi + 1]

        nxt_start = nxt_group[0].start_ms

        # Determine if octave change is needed between groups
        cur_pitches = [n.pitch for n in cur_group]
        nxt_pitches = [n.pitch for n in nxt_group]
        cur_octs = set(_note_octave(p) for p in cur_pitches)
        nxt_octs = set(_note_octave(p) for p in nxt_pitches)
        needs_octave_change = cur_octs != nxt_octs

        min_gap = GW2_MIN_NOTE_DELAY_MS
        if needs_octave_change:
            min_gap = GW2_MIN_NOTE_DELAY_MS + GW2_OCTAVE_SWAP_DELAY_MS

        # Trim each note in current group so it ends min_gap ms before next group
        for note in cur_group:
            note_end = note.start_ms + note.duration_ms
            gap = nxt_start - note_end
            if gap < min_gap:
                new_dur = nxt_start - note.start_ms - min_gap
                if new_dur < 10:
                    new_dur = 10  # floor: don't shrink below 10ms
                if new_dur != note.duration_ms:
                    note.duration_ms = new_dur
                    fixes += 1

        # If trimming to floor still leaves insufficient gap, shift next group forward
        latest_end = max(n.start_ms + n.duration_ms for n in cur_group)
        actual_gap = nxt_start - latest_end
        if actual_gap < min_gap:
            shift = min_gap - actual_gap
            for note in nxt_group:
                note.start_ms += shift
            fixes += 1

    return fixes

def analyze_song(notes, tracks, inst_lo, inst_hi):
    """Comprehensive GW2 playback analysis across all tracks.

    Args:
        notes: list of PianoRollNote
        tracks: list of track info dicts (with 'melody', 'name', etc.)
        inst_lo: instrument low MIDI pitch
        inst_hi: instrument high MIDI pitch

    Returns a report dict with per-track issues and suggested fixes,
    or None if no active notes.
    """
    base_pitch = inst_lo + 12  # mid octave root
    active = [n for n in notes if not n.deleted and not n.simplified]
    if not active:
        return None

    melody_idxs = set(i for i, t in enumerate(tracks) if t.get('melody', False))
    visible_idxs = set(i for i, t in enumerate(tracks) if t.get('visible', True))
    CHORD_WIN = 5  # ms

    # Build melody pitch-class lookup by time for bass duplicate detection
    melody_notes = sorted(
        [n for n in active if n.track in melody_idxs],
        key=lambda n: n.start_ms)
    mel_lookup = []  # (start_ms, set of (pitch_class, pitch))
    mi = 0
    while mi < len(melody_notes):
        t0 = melody_notes[mi].start_ms
        pcs = set()
        mj = mi
        while mj < len(melody_notes) and melody_notes[mj].start_ms - t0 <= CHORD_WIN:
            pcs.add((melody_notes[mj].pitch % 12, melody_notes[mj].pitch))
            mj += 1
        mel_lookup.append((t0, pcs))
        mi = mj

    track_reports = []

    for tidx, tinfo in enumerate(tracks):
        if tidx not in visible_idxs:
            continue
        tnotes = sorted(
            [n for n in active if n.track == tidx],
            key=lambda n: n.start_ms)
        if not tnotes:
            continue

        is_mel = tinfo.get('melody', False)
        r = {
            'index': tidx,
            'name': tinfo['name'],
            'is_melody': is_mel,
            'note_count': len(tnotes),
            'out_of_range': 0,
            'octave_switches': 0,
            'rapid_switches': 0,
            'bass_dupes': 0,
            'dense_chords': 0,
            'tight_notes': 0,
            'fixes': [],
        }

        # Out of range
        for n in tnotes:
            if n.pitch < inst_lo or n.pitch > inst_hi:
                r['out_of_range'] += 1

        # Octave switches and rapid switches
        prev_oct = None
        prev_end = 0
        for n in tnotes:
            o = _note_octave(n.pitch, base_pitch)
            if prev_oct is not None and o != prev_oct:
                r['octave_switches'] += 1
                gap = n.start_ms - prev_end
                if gap < GW2_MIN_NOTE_DELAY_MS + GW2_OCTAVE_SWAP_DELAY_MS:
                    r['rapid_switches'] += 1
            prev_oct = o
            prev_end = n.start_ms + n.duration_ms

        # Bass duplicates (non-melody tracks only)
        if not is_mel and mel_lookup:
            bi = 0
            for n in tnotes:
                while bi < len(mel_lookup) - 1 and mel_lookup[bi + 1][0] <= n.start_ms + CHORD_WIN:
                    bi += 1
                for bk in range(max(0, bi - 1), min(len(mel_lookup), bi + 2)):
                    bstart, bpcs = mel_lookup[bk]
                    if abs(bstart - n.start_ms) <= CHORD_WIN:
                        pc = n.pitch % 12
                        if any(pc == mpc and n.pitch < mp for mpc, mp in bpcs):
                            r['bass_dupes'] += 1
                            break

        # Dense chords (>2 simultaneous notes in one track)
        ci = 0
        while ci < len(tnotes):
            grp = 1
            cj = ci + 1
            while cj < len(tnotes) and tnotes[cj].start_ms - tnotes[ci].start_ms < CHORD_WIN:
                grp += 1
                cj += 1
            if grp > 2:
                r['dense_chords'] += 1
            ci = cj

        # Tight timing between sequential notes
        for i in range(len(tnotes) - 1):
            gap = tnotes[i + 1].start_ms - (tnotes[i].start_ms + tnotes[i].duration_ms)
            co = _note_octave(tnotes[i].pitch, base_pitch)
            no = _note_octave(tnotes[i + 1].pitch, base_pitch)
            min_gap = GW2_MIN_NOTE_DELAY_MS + (GW2_OCTAVE_SWAP_DELAY_MS if co != no else 0)
            if 0 <= gap < min_gap:
                r['tight_notes'] += 1

        # Suggest fixes based on detected issues
        # Melody: only suggest octave shift if a large portion is OOR.
        # Small OOR counts are acceptable — shifting creates octave changes
        # across all tracks which is worse than a few clamped notes.
        is_preserved = tinfo.get('preserve', False)
        oor_threshold = 0.2 if is_mel else 0.0
        if r['out_of_range'] > len(tnotes) * oor_threshold:
            if r['out_of_range'] > 0:
                r['fixes'].append('smart_octave')
            if not is_mel and r['out_of_range'] > len(tnotes) * 0.2:
                r['fixes'].append('clamp')
        # Preserved tracks: always clamp any remaining OOR (every effort to keep)
        if is_preserved and r['out_of_range'] > 0 and 'clamp' not in r['fixes']:
            r['fixes'].append('clamp')
        if r['bass_dupes'] > 0 and not is_preserved:
            r['fixes'].append('debass')
        if not is_mel and not is_preserved and r['octave_switches'] > len(tnotes) * 0.25:
            r['fixes'].append('simplify')
        if r['tight_notes'] > 0:
            r['fixes'].append('timing')

        track_reports.append(r)

    # Cross-track: simultaneous notes in different octaves
    all_sorted = sorted(active, key=lambda n: n.start_ms)
    cross_oct = 0
    ci = 0
    while ci < len(all_sorted):
        grp = [all_sorted[ci]]
        cj = ci + 1
        while cj < len(all_sorted) and all_sorted[cj].start_ms - all_sorted[ci].start_ms < CHORD_WIN:
            grp.append(all_sorted[cj])
            cj += 1
        octs = set(_note_octave(n.pitch, base_pitch) for n in grp)
        if len(octs) > 1:
            cross_oct += 1
        ci = cj

    return {
        'tracks': track_reports,
        'cross_octave_conflicts': cross_oct,
        'total_notes': len(active),
    }

class PianoRollNote:
    __slots__ = ('start_ms', 'duration_ms', 'pitch', 'velocity', 'track', 'selected', 'deleted', 'warning', 'simplified', 'simplified_manual')
    def __init__(self, start_ms, duration_ms, pitch, velocity, track):
        self.start_ms = start_ms
        self.duration_ms = duration_ms
        self.pitch = pitch
        self.velocity = velocity
        self.track = track
        self.selected = False
        self.deleted = False
        self.warning = ''  # '', 'overlap', 'too_close', 'octave_tight'
        self.simplified = False  # True if note would be removed by simplification
        self.simplified_manual = None  # None=auto, True=force simplified, False=force kept


def extract_notes_with_duration(mid):
    """Extract notes with duration from MIDI. Returns (list of PianoRollNote, bpm)."""
    tpb = mid.ticks_per_beat

    # Build tempo map from all tracks
    tempo_map = []
    for track in mid.tracks:
        tick = 0
        for msg in track:
            tick += msg.time
            if msg.type == 'set_tempo':
                tempo_map.append((tick, msg.tempo))
    if not tempo_map:
        tempo_map = [(0, 500000)]
    tempo_map.sort()

    def ticks_to_ms(target_tick):
        ms = 0.0
        current_tick = 0
        current_tempo = tempo_map[0][1]
        tempo_idx = 0
        while current_tick < target_tick:
            next_tempo_tick = tempo_map[tempo_idx + 1][0] if tempo_idx + 1 < len(tempo_map) else target_tick
            end_tick = min(next_tempo_tick, target_tick)
            delta_ticks = end_tick - current_tick
            ms += (delta_ticks / tpb) * (current_tempo / 1000.0)
            current_tick = end_tick
            if current_tick >= next_tempo_tick and tempo_idx + 1 < len(tempo_map):
                tempo_idx += 1
                current_tempo = tempo_map[tempo_idx][1]
        return ms

    initial_bpm = round(60000000 / tempo_map[0][1])

    notes = []
    for track_idx, track in enumerate(mid.tracks):
        active = {}
        abs_tick = 0
        for msg in track:
            abs_tick += msg.time
            if msg.type == 'note_on' and msg.velocity > 0:
                key = (msg.note, getattr(msg, 'channel', 0))
                active[key] = (abs_tick, msg.velocity)
            elif msg.type == 'note_off' or (msg.type == 'note_on' and msg.velocity == 0):
                key = (msg.note, getattr(msg, 'channel', 0))
                if key in active:
                    start_tick, vel = active.pop(key)
                    start_ms = ticks_to_ms(start_tick)
                    end_ms = ticks_to_ms(abs_tick)
                    dur = max(10, end_ms - start_ms)
                    notes.append(PianoRollNote(start_ms, dur, msg.note, vel, track_idx))
        # Flush lingering notes
        for (pitch, ch), (start_tick, vel) in active.items():
            start_ms = ticks_to_ms(start_tick)
            end_ms = ticks_to_ms(abs_tick)
            dur = max(10, end_ms - start_ms)
            notes.append(PianoRollNote(start_ms, dur, pitch, vel, track_idx))

    notes.sort(key=lambda n: n.start_ms)
    return notes, initial_bpm




def parse_ahk_to_notes(ahk_path):
    """Parse an AHK script and return (list of PianoRollNote, bpm, title, author).

    Resolves octave changes to absolute MIDI pitches.
    Key mapping (GW2 default):
      Numpad 1-8 -> C,D,E,F,G,A,B,C'  (semitone offsets: 0,2,4,5,7,9,11,12)
      F1-F5      -> C#,D#,F#,G#,A#     (semitone offsets: 1,3,6,8,10)
      Numpad 9   -> octave down
      Numpad 0   -> octave up
    Mid octave base = MIDI 60 (C4).
    """
    import re as _re

    KEY_SEMITONES = {
        1: 0, 2: 2, 3: 4, 4: 5, 5: 7, 6: 9, 7: 11, 8: 12,  # natural notes
        9: 1, 10: 3, 11: 6, 12: 8, 13: 10,  # sharps (F1-F5 mapped to 9-13)
    }
    OCTAVE_UP = 100
    OCTAVE_DOWN = 101
    MID_BASE = 60  # MIDI note for C4

    with open(ahk_path, 'r', errors='replace') as f:
        lines = f.readlines()

    title = os.path.splitext(os.path.basename(ahk_path))[0]
    title = title.replace('_', ' ')
    author = ''

    # Parse brace content to key code
    def parse_brace(content):
        c = content.strip().lower()
        # Skip "up" events
        if c.endswith(' up'):
            return -1
        # Strip " down"
        if c.endswith(' down'):
            c = c[:-5].strip()
        # Numpad keys
        m = _re.match(r'numpad(\d)', c)
        if m:
            digit = int(m.group(1))
            if digit == 9:
                return OCTAVE_DOWN
            if digit == 0:
                return OCTAVE_UP
            return digit  # 1-8
        # Bare digit
        if len(c) == 1 and c.isdigit():
            digit = int(c)
            if digit == 9:
                return OCTAVE_DOWN
            if digit == 0:
                return OCTAVE_UP
            return digit
        # F-keys (sharps)
        m = _re.match(r'f(\d+)', c)
        if m:
            fnum = int(m.group(1))
            if 1 <= fnum <= 5:
                return 8 + fnum  # F1->9, F2->10, F3->11, F4->12, F5->13
        return -1

    # First pass: collect events as (type, value) where type is 'key' or 'sleep'
    events = []
    for line in lines:
        trimmed = line.strip()
        if not trimmed:
            continue
        # Metadata comments
        if trimmed.startswith(';') or trimmed.startswith('#'):
            meta = trimmed.lstrip(';#').strip()
            if meta.lower().startswith('title:'):
                title = meta[6:].strip()
            elif meta.lower().startswith('author:'):
                author = meta[7:].strip()
            continue

        lower = trimmed.lower()

        # SendInput / Send lines
        if 'sendinput' in lower or ('send' in lower and 'sleep' not in lower):
            # Try brace-wrapped keys first: SendInput {key}
            pos = 0
            found_brace = False
            while pos < len(trimmed):
                bstart = trimmed.find('{', pos)
                if bstart == -1:
                    break
                bend = trimmed.find('}', bstart + 1)
                if bend == -1:
                    break
                key = parse_brace(trimmed[bstart+1:bend])
                if key >= 0:
                    events.append(('key', key))
                    found_brace = True
                pos = bend + 1
            # Also handle bare keys: SendInput 0, SendInput 9, SendInput 3
            if not found_brace:
                m = _re.search(r'(?:sendinput|send)\s+(\S+)', trimmed, _re.IGNORECASE)
                if m:
                    key = parse_brace(m.group(1))
                    if key >= 0:
                        events.append(('key', key))
            continue

        # Sleep lines
        if 'sleep' in lower:
            m = _re.search(r'sleep[\s,]+(\d+)', lower)
            if m:
                events.append(('sleep', int(m.group(1))))
            continue

    # Second pass: resolve to PianoRollNote objects
    notes = []
    current_time_ms = 0.0
    octave_offset = 0  # 0 = Mid, +1 = High, -1 = Low
    pending_keys = []  # keys accumulated before a sleep

    def flush_pending(sleep_ms):
        nonlocal current_time_ms
        if not pending_keys:
            # Pure rest
            current_time_ms += sleep_ms
            return
        # All pending keys form a chord at current_time_ms
        # Duration = sleep_ms (or minimum 80ms for visual clarity)
        dur = max(80, sleep_ms) if sleep_ms > 0 else 150
        for key_code in pending_keys:
            semi = KEY_SEMITONES.get(key_code)
            if semi is not None:
                midi_pitch = MID_BASE + (octave_offset * 12) + semi
                midi_pitch = max(0, min(127, midi_pitch))
                notes.append(PianoRollNote(
                    start_ms=current_time_ms,
                    duration_ms=dur,
                    pitch=midi_pitch,
                    velocity=100,
                    track=0
                ))
        pending_keys.clear()
        current_time_ms += sleep_ms

    for etype, evalue in events:
        if etype == 'key':
            if evalue == OCTAVE_UP:
                octave_offset = min(1, octave_offset + 1)
            elif evalue == OCTAVE_DOWN:
                octave_offset = max(-1, octave_offset - 1)
            else:
                pending_keys.append(evalue)
        elif etype == 'sleep':
            flush_pending(evalue)

    # Flush any remaining
    flush_pending(0)

    # Estimate BPM from average sleep time
    sleep_times = [v for t, v in events if t == 'sleep' and v > 0]
    if sleep_times:
        avg_ms = sum(sleep_times) / len(sleep_times)
        bpm = max(30, min(300, 60000.0 / (avg_ms * 2)))  # rough estimate
    else:
        bpm = 120

    return notes, bpm, title, author

class PianoRollWidget(QWidget):
    NOTE_HEIGHT = 14
    PIANO_WIDTH = 50
    RULER_HEIGHT = 24
    SCROLLBAR_SIZE = 14
    GW2_PITCH_MIN = 48   # C3 (Low octave bottom)
    GW2_PITCH_MAX = 83   # B5 (High octave top)

    notesChanged = pyqtSignal()
    selectionChanged = pyqtSignal()  # emitted when note selection or range changes
    pianoRollContextMenu = pyqtSignal(object)  # emitted with QPoint global pos on right-click
    scrollChanged = pyqtSignal()  # emitted when scroll position changes

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMinimumSize(400, 542)
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self._notes = []
        self._tracks = []
        self._px_per_ms = 0.5
        self._scroll_x = 0.0
        self._scroll_y = 0
        self._pitch_min = 48
        self._pitch_max = 84
        self._total_ms = 0.0
        self._selecting = False
        self._sel_start = None
        self._sel_end = None
        self._bpm = 120
        self._tempo_scale = 1.0  # 1.0 = original speed, 0.5 = double speed, 2.0 = half speed
        self._cursor_ms = -1.0  # playback cursor position (-1 = hidden)
        self._range_start_ms = -1.0  # ruler range selection start (-1 = none)
        self._range_end_ms = -1.0    # ruler range selection end
        self._dragging_range = False

        # Editing state
        self._edit_mode = 'select'  # 'select' or 'draw'
        self._drawing = False       # currently drawing a new note
        self._draw_start_ms = 0.0
        self._draw_pitch = 60
        self._draw_end_ms = 0.0
        self._dragging_notes = False  # dragging selected notes
        self._drag_start_ms = 0.0
        self._drag_start_pitch = 0
        self._drag_offset_ms = 0.0
        self._drag_offset_pitch = 0
        self._resizing = False       # resizing a note's right edge
        self._resize_note = None
        self._resize_orig_dur = 0.0
        self._resize_start_x = 0.0
        self._snap_ms = 0.0  # snap grid in ms (0 = no snap)
        self._gw2_only = False  # True = clamp to GW2 range (New/AHK)

        # Undo / redo stacks (list of note snapshots)
        self._undo_stack = []
        self._redo_stack = []
        self._undo_max = 50

        # Clipboard
        self._clipboard = []  # list of (relative_ms, duration_ms, pitch, velocity, track)

        # Scrollbars (horizontal hidden — minimap handles horizontal navigation)
        self._hscroll = QScrollBar(Qt.Orientation.Horizontal, self)
        self._hscroll.setFixedHeight(self.SCROLLBAR_SIZE)
        self._hscroll.valueChanged.connect(self._on_hscroll)
        self._hscroll.hide()
        self._vscroll = QScrollBar(Qt.Orientation.Vertical, self)
        self._vscroll.setFixedWidth(self.SCROLLBAR_SIZE)
        self._vscroll.valueChanged.connect(self._on_vscroll)
        self._updating_scrollbars = False

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._layout_scrollbars()
        self._update_scrollbars()

    def _layout_scrollbars(self):
        w, h = self.width(), self.height()
        self._hscroll.setGeometry(self.PIANO_WIDTH, h - self.SCROLLBAR_SIZE,
                                   w - self.PIANO_WIDTH - self.SCROLLBAR_SIZE, self.SCROLLBAR_SIZE)
        self._vscroll.setGeometry(w - self.SCROLLBAR_SIZE, self.RULER_HEIGHT,
                                   self.SCROLLBAR_SIZE, h - self.RULER_HEIGHT - self.SCROLLBAR_SIZE)

    def _update_scrollbars(self):
        self._updating_scrollbars = True
        # Horizontal
        view_w = max(1, self.width() - self.PIANO_WIDTH - self.SCROLLBAR_SIZE)
        total_px = self._total_ms * self._px_per_ms
        if total_px > view_w:
            self._hscroll.setRange(0, int(total_px - view_w))
            self._hscroll.setPageStep(int(view_w))
            self._hscroll.setValue(int(self._scroll_x * self._px_per_ms))
        else:
            self._hscroll.setRange(0, 0)
        self._hscroll.setVisible(False)
        # Vertical
        pitch_rows = self._pitch_max - self._pitch_min + 1
        view_h = self.height() - self.RULER_HEIGHT - self.SCROLLBAR_SIZE
        total_row_px = pitch_rows * self.NOTE_HEIGHT
        if total_row_px > view_h:
            max_scroll = pitch_rows - view_h / self.NOTE_HEIGHT
            self._vscroll.setRange(0, int(max_scroll * 100))
            self._vscroll.setPageStep(int(view_h / self.NOTE_HEIGHT * 100))
            self._vscroll.setValue(int(self._scroll_y * 100))
            self._vscroll.setVisible(True)
        else:
            self._vscroll.setRange(0, 0)
            self._vscroll.setVisible(False)
        self._updating_scrollbars = False
        self.scrollChanged.emit()

    def _on_hscroll(self, value):
        if self._updating_scrollbars:
            return
        self._scroll_x = value / self._px_per_ms
        self.update()

    def _on_vscroll(self, value):
        if self._updating_scrollbars:
            return
        self._scroll_y = value / 100.0
        self.update()

    def setCursorMs(self, ms):
        self._cursor_ms = ms
        self.update()

    def setNotes(self, notes, tracks, bpm=120):
        self._notes = notes
        self._tracks = tracks
        self._bpm = bpm
        # Snap to 1/4 beat (16th note)
        if bpm > 0:
            self._snap_ms = 60000.0 / bpm / 4
        else:
            self._snap_ms = 0
        if self._gw2_only:
            # In GW2 mode, keep the pitch range locked to the instrument
            self._pitch_min = self.GW2_PITCH_MIN
            self._pitch_max = self.GW2_PITCH_MAX
            if notes:
                self._total_ms = max(n.start_ms + n.duration_ms for n in notes) + 1000
            else:
                self._total_ms = 0
        elif notes:
            pitches = [n.pitch for n in notes]
            self._pitch_min = max(0, min(min(pitches) - 4, self.GW2_PITCH_MIN - 2))
            self._pitch_max = min(127, max(max(pitches) + 4, self.GW2_PITCH_MAX + 2))
            self._total_ms = max(n.start_ms + n.duration_ms for n in notes) + 1000
        else:
            self._pitch_min = 48
            self._pitch_max = 84
            self._total_ms = 0
        self._scroll_x = 0.0
        self._scroll_y = 0
        # Zoom to show ~8 beats at a readable scale
        if bpm > 0:
            beat_ms = 60000.0 / bpm
            target_beats_visible = 16
            view_w = max(1, self.width() - self.PIANO_WIDTH - self.SCROLLBAR_SIZE)
            self._px_per_ms = view_w / (target_beats_visible * beat_ms)
            self._px_per_ms = max(0.05, min(5.0, self._px_per_ms))
        # Center vertically on the note range
        pitch_rows = self._pitch_max - self._pitch_min + 1
        view_rows = (self.height() - self.RULER_HEIGHT - self.SCROLLBAR_SIZE) / self.NOTE_HEIGHT
        if pitch_rows > view_rows:
            self._scroll_y = max(0, (pitch_rows - view_rows) / 2)
        self._update_scrollbars()
        self.update()

    def _track_offsets(self):
        """Return dict of track_idx -> time_offset_ms."""
        return {i: t.get('time_offset_ms', 0) for i, t in enumerate(self._tracks)}

    def getActiveNotes(self):
        """Return list of (time_ms, pitch, velocity, is_melody) for non-deleted, visible-track, non-simplified notes."""
        visible = set()
        melody = set()
        for i, t in enumerate(self._tracks):
            if t.get('visible', True):
                visible.add(i)
            if t.get('melody', False):
                melody.add(i)
        offsets = self._track_offsets()
        ts = self._tempo_scale
        result = []
        for n in self._notes:
            if not n.deleted and not n.simplified and n.track in visible:
                t_ms = max(0, n.start_ms + offsets.get(n.track, 0)) * ts
                result.append((t_ms, n.pitch, n.velocity, n.track in melody))
        result.sort(key=lambda x: x[0])
        return result

    def getSelectedNotes(self):
        """Return list of (start_ms, duration_ms, pitch, velocity, is_melody) for selected, non-simplified notes."""
        melody = set(i for i, t in enumerate(self._tracks) if t.get('melody', False))
        offsets = self._track_offsets()
        ts = self._tempo_scale
        result = []
        for n in self._notes:
            if n.selected and not n.deleted and not n.simplified:
                if n.track < len(self._tracks) and not self._tracks[n.track].get('visible', True):
                    continue
                t_ms = max(0, n.start_ms + offsets.get(n.track, 0)) * ts
                result.append((t_ms, n.duration_ms * ts, n.pitch, n.velocity, n.track in melody))
        result.sort(key=lambda x: x[0])
        return result

    def getRangeNotes(self):
        """Return visible notes within the ruler-selected time range."""
        if self._range_start_ms < 0 or self._range_end_ms < 0:
            return []
        lo = min(self._range_start_ms, self._range_end_ms)
        hi = max(self._range_start_ms, self._range_end_ms)
        visible = set()
        melody = set()
        for i, t in enumerate(self._tracks):
            if t.get('visible', True):
                visible.add(i)
            if t.get('melody', False):
                melody.add(i)
        offsets = self._track_offsets()
        ts = self._tempo_scale
        result = []
        for n in self._notes:
            if n.deleted or n.simplified or n.track not in visible:
                continue
            t_ms = max(0, n.start_ms + offsets.get(n.track, 0))
            # Include note if it overlaps the range (compare in original ms)
            if t_ms + n.duration_ms >= lo and t_ms <= hi:
                result.append((t_ms * ts, n.duration_ms * ts, n.pitch, n.velocity, n.track in melody))
        result.sort(key=lambda x: x[0])
        return result

    def _filter_in_range(self, notes):
        """Remove notes outside the GW2 instrument pitch range."""
        lo, hi = self.GW2_PITCH_MIN, self.GW2_PITCH_MAX
        return [n for n in notes if lo <= n[2] <= hi]

    def getPlaybackNotes(self):
        """Return the best set of notes for playback:
        1. Selected notes if any
        2. Range notes if ruler range is set
        3. All visible notes
        Returns (notes_list, offset_ms) where offset_ms is the start time offset.
        Notes are 5-tuples: (start_ms, duration_ms, pitch, velocity, is_melody).
        Notes outside the GW2 instrument range are excluded."""
        selected = self.getSelectedNotes()
        if selected:
            offset = selected[0][0] if selected else 0
            adjusted = [(s - offset, d, p, v, m) for s, d, p, v, m in selected]
            return self._filter_in_range(adjusted), offset

        ranged = self.getRangeNotes()
        if ranged:
            offset = min(self._range_start_ms, self._range_end_ms) * self._tempo_scale
            adjusted = [(max(0, s - offset), d, p, v, m) for s, d, p, v, m in ranged]
            return self._filter_in_range(adjusted), offset

        # All visible
        melody = set(i for i, t in enumerate(self._tracks) if t.get('melody', False))
        visible = set(i for i, t in enumerate(self._tracks) if t.get('visible', True))
        offsets = self._track_offsets()
        ts = self._tempo_scale
        all_notes = []
        for n in self._notes:
            if not n.deleted and not n.simplified and n.track in visible:
                t_ms = max(0, n.start_ms + offsets.get(n.track, 0)) * ts
                all_notes.append((t_ms, n.duration_ms * ts, n.pitch, n.velocity, n.track in melody))
        all_notes.sort(key=lambda x: x[0])
        return self._filter_in_range(all_notes), 0

    def getPlaybackMode(self):
        """Return 'selection', 'range', or 'all'."""
        selected = [n for n in self._notes if n.selected and not n.deleted]
        if selected:
            return 'selection'
        if self._range_start_ms >= 0 and self._range_end_ms >= 0:
            lo = min(self._range_start_ms, self._range_end_ms)
            hi = max(self._range_start_ms, self._range_end_ms)
            if hi - lo > 10:
                return 'range'
        return 'all'

    def clearRange(self):
        self._range_start_ms = -1.0
        self._range_end_ms = -1.0
        self.update()
        self.selectionChanged.emit()

    def setTrackVisible(self, track_idx, visible):
        if track_idx < len(self._tracks):
            self._tracks[track_idx]['visible'] = visible
            self.updateSimplifiedNotes()
            self.update()

    def updateSimplifiedNotes(self):
        """Mark notes that would be removed by per-track simplification.
        For each track with simplify=True, groups that track's notes by time
        and keeps only the highest (treble) and lowest (bass) notes, marking
        the rest as simplified.  Then, if a melody track is set, hide
        lower-octave duplicates from non-melody tracks that share a pitch
        class with a simultaneous melody note.
        Manual overrides are applied afterwards."""
        # Clear auto-computed state
        for n in self._notes:
            n.simplified = False
        visible = set(i for i, t in enumerate(self._tracks) if t.get('visible', True))
        CHORD_WIN = 5
        # ── Pass 1: per-track chord simplification ──
        simplify_tracks = set(i for i, t in enumerate(self._tracks) if t.get('simplify', False))
        for tidx in simplify_tracks:
            if tidx not in visible:
                continue
            track_notes = [n for n in self._notes if not n.deleted and n.track == tidx]
            track_notes.sort(key=lambda n: n.start_ms)
            if not track_notes:
                continue
            i = 0
            while i < len(track_notes):
                group = [track_notes[i]]
                j = i + 1
                while j < len(track_notes) and track_notes[j].start_ms - group[0].start_ms <= CHORD_WIN:
                    group.append(track_notes[j])
                    j += 1
                if len(group) >= 3:
                    pitches = [n.pitch for n in group]
                    lo, hi = min(pitches), max(pitches)
                    for n in group:
                        if n.pitch != lo and n.pitch != hi:
                            n.simplified = True
                i = j
        # ── Pass 2: melody priority — hide lower-octave duplicates ──
        melody_tracks = set(i for i, t in enumerate(self._tracks) if t.get('melody', False))
        if melody_tracks:
            offsets = self._track_offsets()
            # Collect melody notes grouped by effective time (with offset)
            melody_notes = [n for n in self._notes
                            if not n.deleted and not n.simplified
                            and n.track in melody_tracks and n.track in visible]
            melody_notes.sort(key=lambda n: n.start_ms + offsets.get(n.track, 0))
            # Build time buckets: list of (bucket_start_ms, set of (pitch_class, pitch))
            buckets = []
            i = 0
            while i < len(melody_notes):
                bucket_start = melody_notes[i].start_ms + offsets.get(melody_notes[i].track, 0)
                pitches = set()
                j = i
                while j < len(melody_notes) and (melody_notes[j].start_ms + offsets.get(melody_notes[j].track, 0)) - bucket_start <= CHORD_WIN:
                    pitches.add((melody_notes[j].pitch % 12, melody_notes[j].pitch))
                    j += 1
                buckets.append((bucket_start, pitches))
                i = j
            # Scan non-melody visible notes
            debass_tracks = set(i for i, t in enumerate(self._tracks) if t.get('debass', False))
            non_melody = [n for n in self._notes
                          if not n.deleted and not n.simplified
                          and n.track in debass_tracks and n.track in visible]
            non_melody.sort(key=lambda n: n.start_ms + offsets.get(n.track, 0))
            bi = 0
            for n in non_melody:
                n_ms = n.start_ms + offsets.get(n.track, 0)
                # Advance bucket index past buckets that are too early
                while bi > 0 and bi < len(buckets) and buckets[bi][0] > n_ms + CHORD_WIN:
                    bi -= 1
                while bi < len(buckets) and buckets[bi][0] < n_ms - CHORD_WIN:
                    bi += 1
                # Check nearby buckets for pitch-class collision
                for bk in range(bi, len(buckets)):
                    bstart, bpitches = buckets[bk]
                    if bstart > n_ms + CHORD_WIN:
                        break
                    if abs(bstart - n_ms) <= CHORD_WIN:
                        pc = n.pitch % 12
                        for mel_pc, mel_pitch in bpitches:
                            if pc == mel_pc and n.pitch < mel_pitch:
                                n.simplified = True
                                break
                    if n.simplified:
                        break
        # Apply manual overrides
        for n in self._notes:
            if n.simplified_manual is True:
                n.simplified = True
            elif n.simplified_manual is False:
                n.simplified = False

    def _pitch_to_y(self, pitch):
        row = self._pitch_max - pitch
        return self.RULER_HEIGHT + (row - self._scroll_y) * self.NOTE_HEIGHT

    def _ms_to_x(self, ms):
        return self.PIANO_WIDTH + (ms - self._scroll_x) * self._px_per_ms

    def _x_to_ms(self, x):
        return (x - self.PIANO_WIDTH) / self._px_per_ms + self._scroll_x

    def _y_to_pitch(self, y):
        row = (y - self.RULER_HEIGHT) / self.NOTE_HEIGHT + self._scroll_y
        return int(self._pitch_max - row)

    def _snap(self, ms):
        if self._snap_ms > 0:
            return round(ms / self._snap_ms) * self._snap_ms
        return ms

    def setInstrumentRange(self, midi_lo, midi_hi):
        """Set pitch range and adjust widget minimum height to show all rows."""
        self._gw2_only = True
        self.GW2_PITCH_MIN = midi_lo
        self.GW2_PITCH_MAX = midi_hi
        self._pitch_min = midi_lo
        self._pitch_max = midi_hi
        pitch_rows = midi_hi - midi_lo + 1
        min_h = pitch_rows * self.NOTE_HEIGHT + self.RULER_HEIGHT + self.SCROLLBAR_SIZE
        self.setMinimumHeight(min_h)
        self._scroll_y = 0
        self._update_scrollbars()
        self.update()

    def _snapshot(self, notes=None):
        """Create a snapshot of the given notes (default: self._notes)."""
        if notes is None:
            notes = self._notes
        snapshot = []
        for n in notes:
            sn = PianoRollNote(n.start_ms, n.duration_ms, n.pitch, n.velocity, n.track)
            sn.selected = n.selected
            sn.deleted = n.deleted
            sn.warning = n.warning
            snapshot.append(sn)
        return snapshot

    def pushUndo(self):
        """Save a snapshot of the current notes for undo."""
        self._undo_stack.append(self._snapshot())
        if len(self._undo_stack) > self._undo_max:
            self._undo_stack.pop(0)
        self._redo_stack.clear()

    def undo(self):
        """Restore the last snapshot from the undo stack."""
        if not self._undo_stack:
            return False
        self._redo_stack.append(self._snapshot())
        if len(self._redo_stack) > self._undo_max:
            self._redo_stack.pop(0)
        self._notes[:] = self._undo_stack.pop()
        self.notesChanged.emit()
        self.update()
        return True

    def redo(self):
        """Re-apply the last undone change."""
        if not self._redo_stack:
            return False
        self._undo_stack.append(self._snapshot())
        if len(self._undo_stack) > self._undo_max:
            self._undo_stack.pop(0)
        self._notes[:] = self._redo_stack.pop()
        self.notesChanged.emit()
        self.update()
        return True

    def copySelected(self):
        """Copy selected notes to clipboard."""
        sel = [n for n in self._notes if n.selected and not n.deleted]
        if not sel:
            return
        min_ms = min(n.start_ms for n in sel)
        self._clipboard = [(n.start_ms - min_ms, n.duration_ms, n.pitch, n.velocity, n.track) for n in sel]

    def pasteClipboard(self):
        """Paste clipboard notes at the current cursor or selection start."""
        if not self._clipboard:
            return
        self.pushUndo()
        # Paste at cursor if playing, else at selection start, else at scroll position
        if self._cursor_ms >= 0:
            paste_ms = self._cursor_ms
        elif self._range_start_ms >= 0:
            paste_ms = min(self._range_start_ms, self._range_end_ms)
        else:
            paste_ms = self._scroll_x
        # Deselect all, then add pasted notes as selected
        for n in self._notes:
            n.selected = False
        for rel_ms, dur, pitch, vel, track in self._clipboard:
            nn = PianoRollNote(paste_ms + rel_ms, dur, pitch, vel, track)
            nn.selected = True
            self._notes.append(nn)
        self.notesChanged.emit()
        self.update()

    def zoomIn(self):
        """Zoom in by 30%, centered on the viewport middle."""
        mid_ms = self._scroll_x + (self.width() - self.PIANO_WIDTH - self.SCROLLBAR_SIZE) / 2 / self._px_per_ms
        self._px_per_ms = min(8.0, self._px_per_ms * 1.25)
        new_mid_ms = self._scroll_x + (self.width() - self.PIANO_WIDTH - self.SCROLLBAR_SIZE) / 2 / self._px_per_ms
        self._scroll_x += mid_ms - new_mid_ms
        self._scroll_x = max(0, self._scroll_x)
        self._update_scrollbars()
        self.update()

    def zoomOut(self):
        """Zoom out by 30%, centered on the viewport middle."""
        mid_ms = self._scroll_x + (self.width() - self.PIANO_WIDTH - self.SCROLLBAR_SIZE) / 2 / self._px_per_ms
        self._px_per_ms = max(0.02, self._px_per_ms / 1.25)
        new_mid_ms = self._scroll_x + (self.width() - self.PIANO_WIDTH - self.SCROLLBAR_SIZE) / 2 / self._px_per_ms
        self._scroll_x += mid_ms - new_mid_ms
        self._scroll_x = max(0, self._scroll_x)
        self._update_scrollbars()
        self.update()

    def setEditMode(self, mode):
        self._edit_mode = mode
        if mode == 'draw':
            self.setCursor(Qt.CursorShape.CrossCursor)
        else:
            self.setCursor(Qt.CursorShape.ArrowCursor)
        self.update()

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing, False)
        w, h = self.width(), self.height()

        p.fillRect(0, 0, w, h, QColor(30, 30, 30))

        na_x = self.PIANO_WIDTH
        na_y = self.RULER_HEIGHT
        black_semitones = {1, 3, 6, 8, 10}

        # Pitch rows + piano keys
        for pitch in range(self._pitch_min, self._pitch_max + 1):
            y = self._pitch_to_y(pitch)
            if y < na_y - self.NOTE_HEIGHT or y > h:
                continue
            semi = pitch % 12
            is_black = semi in black_semitones
            # Row background
            if is_black:
                p.fillRect(na_x, int(y), w - na_x, self.NOTE_HEIGHT, QColor(25, 25, 25))
            # Out-of-GW2-range shading
            if not self._gw2_only and (pitch < self.GW2_PITCH_MIN or pitch > self.GW2_PITCH_MAX):
                p.fillRect(na_x, int(y), w - na_x, self.NOTE_HEIGHT, QColor(80, 20, 20, 50))
                p.fillRect(0, int(y), self.PIANO_WIDTH - 2, self.NOTE_HEIGHT, QColor(80, 20, 20, 40))
            # Grid line
            if semi == 0:
                p.setPen(QPen(QColor(90, 90, 90), 1, Qt.PenStyle.DotLine))
            else:
                p.setPen(QPen(QColor(42, 42, 42), 1))
            p.drawLine(na_x, int(y), w, int(y))
            # Piano key strip
            if is_black:
                p.fillRect(0, int(y), self.PIANO_WIDTH - 2, self.NOTE_HEIGHT, QColor(40, 40, 40))
            else:
                p.fillRect(0, int(y), self.PIANO_WIDTH - 2, self.NOTE_HEIGHT, QColor(70, 70, 70))
            # Note labels
            note_names = ['C', 'C#', 'D', 'D#', 'E', 'F', 'F#', 'G', 'G#', 'A', 'A#', 'B']
            octave = pitch // 12 - 1
            name = f"{note_names[semi]}{octave}"
            if is_black:
                p.setPen(QColor(140, 140, 140))
            elif semi == 0:
                p.setPen(QColor(220, 220, 220))
            else:
                p.setPen(QColor(180, 180, 180))
            p.setFont(QFont("monospace", 7))
            p.drawText(2, int(y), self.PIANO_WIDTH - 6, self.NOTE_HEIGHT,
                       Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignRight,
                       name)

        # GW2 range boundary lines
        if not self._gw2_only:
            for boundary_pitch in (self.GW2_PITCH_MIN, self.GW2_PITCH_MAX + 1):
                by = self._pitch_to_y(boundary_pitch)
                if boundary_pitch == self.GW2_PITCH_MIN:
                    by += self.NOTE_HEIGHT  # bottom edge of lowest playable row
                if na_y <= by <= h:
                    p.setPen(QPen(QColor(255, 80, 80, 160), 2, Qt.PenStyle.DashLine))
                    p.drawLine(na_x, int(by), w, int(by))

        # Time grid
        if self._bpm > 0:
            beat_ms = 60000.0 / self._bpm
            first_beat = max(0, int(self._scroll_x / beat_ms))
            p.setFont(QFont("monospace", 7))
            ms = first_beat * beat_ms
            while True:
                x = self._ms_to_x(ms)
                if x > w:
                    break
                if x >= na_x:
                    beat_num = int(round(ms / beat_ms))
                    if beat_num % 4 == 0:
                        p.setPen(QPen(QColor(75, 75, 75), 1))
                        p.drawLine(int(x), na_y, int(x), h)
                    else:
                        p.setPen(QPen(QColor(42, 42, 42), 1))
                        p.drawLine(int(x), na_y, int(x), h)
                ms += beat_ms

        # Range highlight
        if self._range_start_ms >= 0 and self._range_end_ms >= 0:
            rx1 = self._ms_to_x(min(self._range_start_ms, self._range_end_ms))
            rx2 = self._ms_to_x(max(self._range_start_ms, self._range_end_ms))
            if rx2 > na_x and rx1 < w:
                rx1 = max(na_x, int(rx1))
                rx2 = min(w, int(rx2))
                p.fillRect(rx1, na_y, rx2 - rx1, h - na_y, QColor(80, 120, 200, 30))
                p.setPen(QPen(QColor(80, 120, 200, 120), 1))
                p.drawLine(rx1, na_y, rx1, h)
                p.drawLine(rx2, na_y, rx2, h)

        # Ruler
        p.fillRect(0, 0, w, self.RULER_HEIGHT, QColor(45, 45, 45))
        if self._bpm > 0:
            beat_ms = 60000.0 / self._bpm
            first_beat = max(0, int(self._scroll_x / beat_ms))
            ms = first_beat * beat_ms
            p.setFont(QFont("monospace", 8))
            while True:
                x = self._ms_to_x(ms)
                if x > w:
                    break
                if x >= na_x:
                    beat_num = int(round(ms / beat_ms))
                    if beat_num % 4 == 0:
                        bar_num = beat_num // 4 + 1
                        p.setPen(QColor(170, 170, 170))
                        p.drawText(int(x) + 3, 2, 50, self.RULER_HEIGHT - 4,
                                   Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter,
                                   str(bar_num))
                        p.setPen(QPen(QColor(100, 100, 100), 1))
                        p.drawLine(int(x), 0, int(x), self.RULER_HEIGHT)
                ms += beat_ms
        # Range highlight on ruler
        if self._range_start_ms >= 0 and self._range_end_ms >= 0:
            rx1 = self._ms_to_x(min(self._range_start_ms, self._range_end_ms))
            rx2 = self._ms_to_x(max(self._range_start_ms, self._range_end_ms))
            if rx2 > na_x and rx1 < w:
                rx1 = max(na_x, int(rx1))
                rx2 = min(w, int(rx2))
                p.fillRect(rx1, 0, rx2 - rx1, self.RULER_HEIGHT, QColor(80, 120, 200, 80))

        p.setPen(QPen(QColor(80, 80, 80), 1))
        p.drawLine(0, self.RULER_HEIGHT, w, self.RULER_HEIGHT)

        # Notes (clip to note area so they don't overlap pitch labels)
        p.setClipRect(na_x, na_y, w - na_x, h - na_y)
        _offsets = self._track_offsets()
        for note in self._notes:
            if note.deleted:
                continue
            if note.track < len(self._tracks) and not self._tracks[note.track].get('visible', True):
                continue
            draw_ms = note.start_ms + _offsets.get(note.track, 0)
            draw_pitch = note.pitch
            if self._dragging_notes and note.selected:
                snap = self._snap_ms if self._snap_ms > 0 else 0
                draw_ms += (self._snap(self._drag_offset_ms) if snap > 0 else self._drag_offset_ms)
                draw_pitch += self._drag_offset_pitch
                draw_ms = max(0, draw_ms)
                draw_pitch = max(self._pitch_min, min(self._pitch_max, draw_pitch))
            x = self._ms_to_x(draw_ms)
            y = self._pitch_to_y(draw_pitch)
            nw = max(2, note.duration_ms * self._px_per_ms)
            if x + nw < na_x or x > w or y + self.NOTE_HEIGHT < na_y or y > h:
                continue
            r, g, b = TRACK_COLORS[note.track % len(TRACK_COLORS)]
            color = QColor(r, g, b)
            out_of_range = not self._gw2_only and (draw_pitch < self.GW2_PITCH_MIN or draw_pitch > self.GW2_PITCH_MAX)
            if out_of_range:
                color = QColor(r // 2, g // 2, b // 2)  # dim out-of-range notes
            if note.simplified:
                color = QColor(80, 30, 30)  # visibly dim simplified-out notes
            if note.selected:
                p.fillRect(int(x), int(y) + 1, int(nw), self.NOTE_HEIGHT - 1, color.lighter(150))
                p.setPen(QPen(QColor(255, 255, 255), 2))
                p.drawRect(int(x), int(y) + 1, int(nw), self.NOTE_HEIGHT - 2)
            else:
                p.fillRect(int(x), int(y) + 1, int(nw), self.NOTE_HEIGHT - 1, color)
                if note.simplified:
                    # Dashed strikethrough to indicate note will be removed
                    p.setPen(QPen(QColor(255, 100, 100, 160), 1, Qt.PenStyle.DashLine))
                    p.drawLine(int(x), int(y) + self.NOTE_HEIGHT // 2, int(x) + int(nw), int(y) + self.NOTE_HEIGHT // 2)
                elif out_of_range:
                    # Draw a subtle strikethrough to indicate unplayable
                    p.setPen(QPen(QColor(255, 60, 60, 100), 1))
                    p.drawLine(int(x), int(y) + self.NOTE_HEIGHT // 2, int(x) + int(nw), int(y) + self.NOTE_HEIGHT // 2)
            # GW2 timing warning overlay
            if note.warning:
                if note.warning == 'chord':
                    warn_color = QColor(100, 140, 255, 80)  # subtle blue for chords (informational)
                elif note.warning == 'octave_tight':
                    warn_color = QColor(255, 160, 0, 140)
                else:  # too_close
                    warn_color = QColor(255, 200, 0, 120)
                p.setPen(QPen(warn_color, 2))
                p.drawRect(int(x), int(y) + 1, int(nw), self.NOTE_HEIGHT - 2)
                # Small warning triangle at left edge
                p.setBrush(QBrush(warn_color))
                p.setPen(Qt.PenStyle.NoPen)
                tx = int(x) + 2
                ty = int(y) + 2
                tri = QPolygonF([QPointF(tx, ty + 8), QPointF(tx + 4, ty), QPointF(tx + 8, ty + 8)])
                p.drawPolygon(tri)
            # Draw note name on the bar
            _nn = ['C', 'C#', 'D', 'D#', 'E', 'F', 'F#', 'G', 'G#', 'A', 'A#', 'B']
            _lbl = f"{_nn[note.pitch % 12]}{note.pitch // 12 - 1}"
            p.setFont(QFont("monospace", 7))
            p.setPen(QColor(0, 0, 0))
            p.drawText(int(x) + 2, int(y) + 1, int(nw) - 2, self.NOTE_HEIGHT - 1,
                       Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft, _lbl)

        # Selection rectangle
        if self._selecting and self._sel_start and self._sel_end:
            sx = min(self._sel_start.x(), self._sel_end.x())
            sy = min(self._sel_start.y(), self._sel_end.y())
            sw = abs(self._sel_end.x() - self._sel_start.x())
            sh = abs(self._sel_end.y() - self._sel_start.y())
            p.setPen(QPen(QColor(255, 255, 255, 150), 1))
            p.setBrush(QBrush(QColor(255, 255, 255, 30)))
            p.drawRect(int(sx), int(sy), int(sw), int(sh))

        # Ghost note while drawing
        if self._drawing:
            gx = self._ms_to_x(min(self._draw_start_ms, self._draw_end_ms))
            gy = self._pitch_to_y(self._draw_pitch)
            gw = abs(self._draw_end_ms - self._draw_start_ms) * self._px_per_ms
            gw = max(4, gw)
            p.fillRect(int(gx), int(gy) + 1, int(gw), self.NOTE_HEIGHT - 1,
                       QColor(100, 200, 100, 120))
            p.setPen(QPen(QColor(100, 200, 100, 200), 1))
            p.drawRect(int(gx), int(gy) + 1, int(gw), self.NOTE_HEIGHT - 2)

        p.setClipping(False)

        # Playback cursor
        if self._cursor_ms >= 0:
            cx = self._ms_to_x(self._cursor_ms)
            if self.PIANO_WIDTH <= cx <= w:
                p.setPen(QPen(QColor(255, 60, 60), 2))
                p.drawLine(int(cx), na_y, int(cx), h)

        # Piano strip border
        p.setPen(QPen(QColor(80, 80, 80), 1))
        p.drawLine(self.PIANO_WIDTH, 0, self.PIANO_WIDTH, h)

        p.end()

    def _note_at(self, pos):
        for note in reversed(self._notes):
            if note.deleted:
                continue
            if note.track < len(self._tracks) and not self._tracks[note.track].get('visible', True):
                continue
            x = self._ms_to_x(note.start_ms)
            y = self._pitch_to_y(note.pitch)
            nw = max(2, note.duration_ms * self._px_per_ms)
            if x <= pos.x() <= x + nw and y <= pos.y() <= y + self.NOTE_HEIGHT:
                return note
        return None

    def _near_right_edge(self, note, pos, threshold=6):
        """Check if pos is near the right edge of a note (for resizing)."""
        x = self._ms_to_x(note.start_ms)
        nw = max(2, note.duration_ms * self._px_per_ms)
        right_x = x + nw
        return abs(pos.x() - right_x) < threshold

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            px, py = event.position().x(), event.position().y()

            # Ruler click = range selection (both modes)
            if py < self.RULER_HEIGHT:
                ms = self._x_to_ms(px)
                self._range_start_ms = max(0, ms)
                self._range_end_ms = self._range_start_ms
                self._dragging_range = True
                for n in self._notes:
                    n.selected = False
                self.update()
                self.selectionChanged.emit()
                return

            # Piano key area - ignore
            if px < self.PIANO_WIDTH:
                return

            note = self._note_at(event.position())
            ctrl = bool(event.modifiers() & Qt.KeyboardModifier.ControlModifier)

            if self._edit_mode == 'draw':
                if note:
                    # In draw mode, clicking existing note = select it for preview/delete
                    if ctrl:
                        note.selected = not note.selected
                    else:
                        for n in self._notes:
                            n.selected = False
                        note.selected = True
                    try:
                        play_preview_note(note.pitch, note.velocity)
                    except Exception:
                        pass
                    self.selectionChanged.emit()
                else:
                    # Draw a new note
                    ms = self._snap(max(0, self._x_to_ms(px)))
                    pitch = self._y_to_pitch(py)
                    if self._gw2_only:
                        pitch = max(self.GW2_PITCH_MIN, min(self.GW2_PITCH_MAX, pitch))
                    else:
                        pitch = max(self._pitch_min, min(self._pitch_max, pitch))
                    self._drawing = True
                    self._draw_start_ms = ms
                    self._draw_end_ms = ms
                    self._draw_pitch = pitch
                    # Clear selection
                    for n in self._notes:
                        n.selected = False
                    self._range_start_ms = -1.0
                    self._range_end_ms = -1.0
                    self.selectionChanged.emit()
                self.update()
                return

            # Select mode
            if note:
                self._range_start_ms = -1.0
                self._range_end_ms = -1.0

                shift = bool(event.modifiers() & Qt.KeyboardModifier.ShiftModifier)
                # Ctrl+Shift+Click: toggle simplified state
                if ctrl and shift:
                    track = self._tracks[note.track] if note.track < len(self._tracks) else {}
                    if track.get('simplify', False):
                        if note.simplified_manual is None:
                            # Override: force opposite of auto state
                            note.simplified_manual = not note.simplified
                        else:
                            # Clear manual override (back to auto)
                            note.simplified_manual = None
                        self.updateSimplifiedNotes()
                        self.update()
                    return

                # Check for right-edge resize
                if self._near_right_edge(note, event.position()):
                    if not note.selected:
                        for n in self._notes:
                            n.selected = False
                        note.selected = True
                    self._resizing = True
                    self._resize_note = note
                    self._resize_orig_dur = note.duration_ms
                    self._resize_start_x = px
                    self.selectionChanged.emit()
                    self.update()
                    return

                if ctrl:
                    note.selected = not note.selected
                else:
                    if not note.selected:
                        for n in self._notes:
                            n.selected = False
                        note.selected = True
                    # Start dragging selected notes
                    self._dragging_notes = True
                    self._drag_start_ms = self._x_to_ms(px)
                    self._drag_start_pitch = self._y_to_pitch(py)
                    self._drag_offset_ms = 0.0
                    self._drag_offset_pitch = 0

                try:
                    play_preview_note(note.pitch, note.velocity)
                except Exception:
                    pass
                self.selectionChanged.emit()
            else:
                self._range_start_ms = -1.0
                self._range_end_ms = -1.0
                if not ctrl:
                    for n in self._notes:
                        n.selected = False
                self._selecting = True
                self._sel_start = event.position()
                self._sel_end = event.position()
                self.selectionChanged.emit()
            self.update()

        elif event.button() == Qt.MouseButton.RightButton:
            px, py = event.position().x(), event.position().y()
            if py < self.RULER_HEIGHT:
                # Right-click on ruler: show trim menu
                ms = self._x_to_ms(px)
                self._show_trim_menu(event.globalPosition().toPoint(), ms)
                return
            # Right-click on note area: if nothing selected, select note under cursor
            if not any(n.selected for n in self._notes):
                note = self._note_at(event.position())
                if note:
                    note.selected = True
                    self.selectionChanged.emit()
                    self.update()
            self.pianoRollContextMenu.emit(event.globalPosition().toPoint())

    def mouseMoveEvent(self, event):
        px, py = event.position().x(), event.position().y()

        if self._drawing:
            ms = self._snap(max(0, self._x_to_ms(px)))
            self._draw_end_ms = ms
            self.update()
            return

        if self._dragging_range:
            ms = self._x_to_ms(px)
            self._range_end_ms = max(0, ms)
            self.update()
            return

        if self._resizing and self._resize_note:
            delta_px = px - self._resize_start_x
            delta_ms = delta_px / self._px_per_ms
            new_dur = max(self._snap_ms if self._snap_ms > 0 else 50,
                          self._resize_orig_dur + delta_ms)
            self._resize_note.duration_ms = self._snap(new_dur) if self._snap_ms > 0 else new_dur
            self.update()
            return

        if self._dragging_notes:
            current_ms = self._x_to_ms(px)
            current_pitch = self._y_to_pitch(py)
            self._drag_offset_ms = current_ms - self._drag_start_ms
            self._drag_offset_pitch = current_pitch - self._drag_start_pitch
            self.update()
            return

        if self._selecting:
            self._sel_end = event.position()
            self.update()
            return

        # Update cursor for resize hint (select mode)
        if self._edit_mode == 'select':
            note = self._note_at(event.position())
            if note and self._near_right_edge(note, event.position()):
                self.setCursor(Qt.CursorShape.SizeHorCursor)
            else:
                self.setCursor(Qt.CursorShape.ArrowCursor)

    def mouseReleaseEvent(self, event):
        if self._drawing:
            # Commit the drawn note
            self.pushUndo()
            start = min(self._draw_start_ms, self._draw_end_ms)
            end = max(self._draw_start_ms, self._draw_end_ms)
            dur = end - start
            if dur < (self._snap_ms if self._snap_ms > 0 else 50):
                # Click without drag = place a note with default duration (1 beat)
                dur = self._snap_ms * 4 if self._snap_ms > 0 else 250
            pitch = self._draw_pitch
            if self._gw2_only:
                pitch = max(self.GW2_PITCH_MIN, min(self.GW2_PITCH_MAX, pitch))
            else:
                pitch = max(self._pitch_min, min(self._pitch_max, pitch))
            note = PianoRollNote(start, dur, pitch, 100, 0)
            self._notes.append(note)
            # Extend total if needed
            if start + dur + 1000 > self._total_ms:
                self._total_ms = start + dur + 2000
                self._update_scrollbars()
            self._drawing = False
            self.notesChanged.emit()
            try:
                play_preview_note(pitch, 100)
            except Exception:
                pass
            self.update()
            return

        if self._resizing:
            self.pushUndo()
            self._resizing = False
            self._resize_note = None
            self.notesChanged.emit()
            self.update()
            return

        if self._dragging_notes:
            # Apply the drag offset to all selected notes
            snap = self._snap_ms if self._snap_ms > 0 else 0
            offset_ms = self._snap(self._drag_offset_ms) if snap > 0 else self._drag_offset_ms
            offset_pitch = self._drag_offset_pitch
            if abs(offset_ms) > 1 or abs(offset_pitch) > 0:
                self.pushUndo()
                for n in self._notes:
                    if n.selected and not n.deleted:
                        n.start_ms = max(0, n.start_ms + offset_ms)
                        pmin = self.GW2_PITCH_MIN if self._gw2_only else self._pitch_min
                        pmax = self.GW2_PITCH_MAX if self._gw2_only else self._pitch_max
                        n.pitch = max(pmin, min(pmax, n.pitch + offset_pitch))
                self.notesChanged.emit()
            self._dragging_notes = False
            self._drag_offset_ms = 0.0
            self._drag_offset_pitch = 0
            self.update()
            return

        if self._dragging_range:
            self._dragging_range = False
            lo = min(self._range_start_ms, self._range_end_ms)
            hi = max(self._range_start_ms, self._range_end_ms)
            if hi - lo < 50:
                self._range_start_ms = -1.0
                self._range_end_ms = -1.0
            self.update()
            self.selectionChanged.emit()
            return

        if self._selecting and self._sel_start and self._sel_end:
            sx = min(self._sel_start.x(), self._sel_end.x())
            sy = min(self._sel_start.y(), self._sel_end.y())
            ex = max(self._sel_start.x(), self._sel_end.x())
            ey = max(self._sel_start.y(), self._sel_end.y())
            if abs(ex - sx) > 3 or abs(ey - sy) > 3:
                for note in self._notes:
                    if note.deleted:
                        continue
                    if note.track < len(self._tracks) and not self._tracks[note.track].get('visible', True):
                        continue
                    nx = self._ms_to_x(note.start_ms)
                    ny = self._pitch_to_y(note.pitch)
                    nw = max(2, note.duration_ms * self._px_per_ms)
                    if nx + nw >= sx and nx <= ex and ny + self.NOTE_HEIGHT >= sy and ny <= ey:
                        note.selected = True
            self._selecting = False
            self._sel_start = None
            self._sel_end = None
            self.update()
            self.selectionChanged.emit()

    def wheelEvent(self, event):
        delta = event.angleDelta().y()
        mods = event.modifiers()
        if mods & Qt.KeyboardModifier.ControlModifier:
            factor = 1.25 if delta > 0 else 1 / 1.25
            mouse_ms = self._x_to_ms(event.position().x())
            self._px_per_ms = max(0.02, min(8.0, self._px_per_ms * factor))
            new_mouse_ms = self._x_to_ms(event.position().x())
            self._scroll_x += mouse_ms - new_mouse_ms
            self._scroll_x = max(0, self._scroll_x)
        elif mods & Qt.KeyboardModifier.ShiftModifier:
            scroll_amount = 300 / self._px_per_ms
            self._scroll_x -= delta / 120 * scroll_amount * 0.3
            self._scroll_x = max(0, min(self._total_ms, self._scroll_x))
        else:
            self._scroll_y -= delta / 120 * 3
            max_scroll = max(0, (self._pitch_max - self._pitch_min + 1) -
                             (self.height() - self.RULER_HEIGHT - self.SCROLLBAR_SIZE) / self.NOTE_HEIGHT)
            self._scroll_y = max(0, min(max_scroll, self._scroll_y))
        self._update_scrollbars()
        self.update()
        event.accept()

    def _compact_after_trim(self, cut_start_ms, cut_end_ms):
        """Shift notes after a deleted region left to close the gap and recalculate duration."""
        gap = cut_end_ms - cut_start_ms
        if gap <= 0:
            return
        for n in self._notes:
            if not n.deleted and n.start_ms >= cut_end_ms:
                n.start_ms -= gap
        active = [n for n in self._notes if not n.deleted]
        if active:
            self._total_ms = max(n.start_ms + n.duration_ms for n in active) + 1000
        else:
            self._total_ms = 0
        self._range_start_ms = -1.0
        self._range_end_ms = -1.0
        self._scroll_x = max(0, min(self._scroll_x, self._total_ms))
        self._update_scrollbars()

    def _show_trim_menu(self, global_pos, ms):
        """Show a context menu to delete/trim notes at a time point or selection range."""
        secs = ms / 1000.0
        menu = QMenu(self)
        before_action = menu.addAction(f"Delete all before {secs:.1f}s")
        after_action = menu.addAction(f"Delete all after {secs:.1f}s")

        # Add selection delete if a range is highlighted
        sel_action = None
        has_range = self._range_start_ms >= 0 and self._range_end_ms >= 0
        if has_range:
            lo = min(self._range_start_ms, self._range_end_ms)
            hi = max(self._range_start_ms, self._range_end_ms)
            if hi - lo > 1:
                sel_action = menu.addAction(f"Delete selection ({lo/1000:.1f}s – {hi/1000:.1f}s)")

        action = menu.exec(global_pos)
        if action is None:
            return
        self.pushUndo()
        count = 0
        if action == before_action:
            for n in self._notes:
                if not n.deleted and n.start_ms < ms:
                    n.deleted = True
                    n.selected = False
                    count += 1
            if count:
                self._compact_after_trim(0, ms)
        elif action == after_action:
            for n in self._notes:
                if not n.deleted and n.start_ms >= ms:
                    n.deleted = True
                    n.selected = False
                    count += 1
            if count:
                active = [n for n in self._notes if not n.deleted]
                if active:
                    self._total_ms = max(n.start_ms + n.duration_ms for n in active) + 1000
                else:
                    self._total_ms = 0
                self._update_scrollbars()
        elif action == sel_action and sel_action is not None:
            for n in self._notes:
                if not n.deleted and n.start_ms >= lo and n.start_ms < hi:
                    n.deleted = True
                    n.selected = False
                    count += 1
            if count:
                self._compact_after_trim(lo, hi)
        if count:
            self.notesChanged.emit()
            self.update()

    def ensureCursorVisible(self):
        """Smooth-scroll so the playback cursor stays at ~30% of the view width."""
        if self._cursor_ms < 0:
            return
        view_w = self.width() - self.PIANO_WIDTH - self.SCROLLBAR_SIZE
        target_x = self._cursor_ms - view_w * 0.3 / self._px_per_ms
        self._scroll_x = max(0, target_x)
        self._update_scrollbars()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_Delete:
            has_selected = any(n.selected for n in self._notes)
            if has_selected:
                self.pushUndo()
            changed = False
            for n in self._notes:
                if n.selected:
                    n.deleted = True
                    n.selected = False
                    changed = True
            if changed:
                self.notesChanged.emit()
                self.update()
        elif event.key() == Qt.Key.Key_A and event.modifiers() & Qt.KeyboardModifier.ControlModifier:
            for n in self._notes:
                if not n.deleted:
                    n.selected = True
            self.update()
        elif event.key() == Qt.Key.Key_Escape:
            for n in self._notes:
                n.selected = False
            self._range_start_ms = -1.0
            self._range_end_ms = -1.0
            self.selectionChanged.emit()
            self.update()
        elif event.key() == Qt.Key.Key_Z and event.modifiers() == (Qt.KeyboardModifier.ControlModifier | Qt.KeyboardModifier.ShiftModifier):
            self.redo()
        elif event.key() == Qt.Key.Key_Z and event.modifiers() & Qt.KeyboardModifier.ControlModifier:
            self.undo()
        elif event.key() == Qt.Key.Key_Y and event.modifiers() & Qt.KeyboardModifier.ControlModifier:
            self.redo()
        elif event.key() == Qt.Key.Key_C and event.modifiers() & Qt.KeyboardModifier.ControlModifier:
            self.copySelected()
        elif event.key() == Qt.Key.Key_V and event.modifiers() & Qt.KeyboardModifier.ControlModifier:
            self.pasteClipboard()
        elif event.key() == Qt.Key.Key_Home:
            self._scroll_x = 0
            self._update_scrollbars()
            self.update()
        elif event.key() == Qt.Key.Key_End:
            view_w = self.width() - self.PIANO_WIDTH - self.SCROLLBAR_SIZE
            self._scroll_x = max(0, self._total_ms - view_w / self._px_per_ms)
            self._update_scrollbars()
            self.update()
        elif event.key() == Qt.Key.Key_Space:
            # Bubble up to MainWindow for playback toggle
            parent = self.parent()
            while parent and not isinstance(parent, QMainWindow):
                parent = parent.parent()
            if parent and hasattr(parent, '_playback_toggle'):
                parent._playback_toggle()
        else:
            super().keyPressEvent(event)


# ── Minimap Widget ────────────────────────────────────────────────────────────

class MinimapWidget(QWidget):
    seekRequested = pyqtSignal(float)  # ms

    def __init__(self, main_window):
        super().__init__()
        self._main = main_window
        self.setMinimumHeight(20)
        self._dragging = False
        self._drag_offset_px = 0.0

    def _viewport_px(self):
        """Return (vx, vw, ms_scale) for the viewport rect in minimap pixel coords."""
        pr = self._main.piano_roll
        max_ms = pr._total_ms
        if max_ms <= 0:
            return None
        w = self.width()
        ms_scale = w / max_ms
        vx = pr._scroll_x * ms_scale
        vw = (pr.width() - pr.PIANO_WIDTH - pr.SCROLLBAR_SIZE) / pr._px_per_ms * ms_scale
        return vx, vw, ms_scale

    def paintEvent(self, event):
        p = QPainter(self)
        p.fillRect(self.rect(), QColor(30, 30, 30))
        mw = self._main
        if not hasattr(mw, '_pr_notes') or not mw._pr_notes:
            p.end()
            return
        active = [n for n in mw._pr_notes if not n.deleted]
        if not active:
            p.end()
            return
        w, h = self.width(), self.height()
        max_ms = mw.piano_roll._total_ms
        if max_ms <= 0:
            p.end()
            return
        ms_scale = w / max_ms
        min_p = mw.piano_roll._pitch_min
        max_p = mw.piano_roll._pitch_max
        p_range = max_p - min_p + 1
        if p_range <= 0:
            p_range = 1
        _offsets = mw.piano_roll._track_offsets()
        for n in active:
            x = max(0, n.start_ms + _offsets.get(n.track, 0)) * ms_scale
            nw = max(1, n.duration_ms * ms_scale)
            y = (max_p - n.pitch) / p_range * h
            nh = max(1, h / p_range)
            cidx = n.track % len(TRACK_COLORS)
            r, g, b = TRACK_COLORS[cidx]
            p.fillRect(int(x), int(y), max(1, int(nw)), max(1, int(nh)), QColor(r, g, b, 180))
        # Draw viewport rectangle
        pr = mw.piano_roll
        vx = pr._scroll_x * ms_scale
        vw = (pr.width() - pr.PIANO_WIDTH - pr.SCROLLBAR_SIZE) / pr._px_per_ms * ms_scale
        p.setPen(QPen(QColor(255, 255, 255, 120), 1))
        p.drawRect(int(vx), 0, int(vw), h - 1)
        # Draw cursor
        if pr._cursor_ms >= 0:
            cx = pr._cursor_ms * ms_scale
            p.setPen(QPen(QColor(255, 80, 80), 1))
            p.drawLine(int(cx), 0, int(cx), h)
        p.end()

    def mousePressEvent(self, event):
        if event.button() != Qt.MouseButton.LeftButton:
            return
        vp = self._viewport_px()
        if vp is None:
            return
        vx, vw, ms_scale = vp
        click_x = event.position().x()
        self._dragging = True
        if vx <= click_x <= vx + vw:
            self._drag_offset_px = click_x - vx
        else:
            self._drag_offset_px = vw / 2
            new_scroll_ms = max(0, (click_x - vw / 2) / ms_scale)
            self.seekRequested.emit(new_scroll_ms)
        event.accept()

    def mouseMoveEvent(self, event):
        if not self._dragging:
            return
        if not (event.buttons() & Qt.MouseButton.LeftButton):
            return
        vp = self._viewport_px()
        if vp is None:
            return
        _, _, ms_scale = vp
        click_x = event.position().x()
        new_scroll_ms = max(0, (click_x - self._drag_offset_px) / ms_scale)
        self.seekRequested.emit(new_scroll_ms)
        event.accept()

    def mouseReleaseEvent(self, event):
        self._dragging = False
        event.accept()


# ── Analysis Dialog ───────────────────────────────────────────────────────────

class AnalysisDialog(QDialog):
    """Dialog showing a comprehensive GW2 playback analysis report with auto-fix option."""
    def __init__(self, report, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Song Analysis")
        self.setMinimumSize(620, 520)
        self.report = report
        self.auto_fix_accepted = False

        layout = QVBoxLayout(self)

        self.text = QTextEdit()
        self.text.setReadOnly(True)
        self.text.setFont(QFont("Consolas", 10))
        layout.addWidget(self.text)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self.fix_btn = QPushButton("Auto-fix All Issues")
        self.fix_btn.setMinimumHeight(32)
        self.fix_btn.setStyleSheet(
            "QPushButton { background-color: #365; border: 2px solid #6b6; border-radius: 4px; "
            "font-weight: bold; padding: 4px 16px; }"
            "QPushButton:hover { background-color: #476; }"
            "QPushButton:disabled { background-color: #333; border-color: #555; color: #777; }")
        self.fix_btn.clicked.connect(self._accept_fix)
        close_btn = QPushButton("Close")
        close_btn.setMinimumHeight(32)
        close_btn.clicked.connect(self.reject)
        btn_row.addWidget(self.fix_btn)
        btn_row.addWidget(close_btn)
        layout.addLayout(btn_row)

        self._build_report()

    def _accept_fix(self):
        self.auto_fix_accepted = True
        self.accept()

    def _build_report(self):
        r = self.report
        lines = []
        lines.append("═══════════════════════════════════════")
        lines.append("         Song Analysis Report")
        lines.append("═══════════════════════════════════════\n")
        lines.append(f"Total active notes: {r['total_notes']}")
        lines.append(f"Cross-octave conflicts: {r['cross_octave_conflicts']}")
        if r['cross_octave_conflicts'] > 0:
            lines.append("  (notes in different octaves at the same time)")

        has_any_fix = False
        for tr in r['tracks']:
            mel = " ♪ MELODY" if tr['is_melody'] else ""
            lines.append(f"\n─── {tr['name']}{mel} ({tr['note_count']} notes) ───")

            issues = []
            if tr['out_of_range']:
                issues.append(f"  ⚠ Out of range: {tr['out_of_range']} notes")
            if tr['octave_switches']:
                rapid = f" ({tr['rapid_switches']} too fast for GW2)" if tr['rapid_switches'] else ""
                issues.append(f"  ⚠ Octave switches: {tr['octave_switches']}{rapid}")
            if tr['bass_dupes']:
                issues.append(f"  ⚠ Bass duplicates: {tr['bass_dupes']} notes duplicate melody at lower octave")
            if tr['dense_chords']:
                issues.append(f"  ⚠ Dense chords: {tr['dense_chords']} groups with >2 simultaneous notes")
            if tr['tight_notes']:
                issues.append(f"  ⚠ Tight timing: {tr['tight_notes']} notes too close together")

            if not issues:
                issues.append("  ✓ No issues")
            lines.extend(issues)

            if tr['fixes']:
                has_any_fix = True
                lines.append("  Suggested fixes:")
                fix_labels = {
                    'smart_octave': '    → Smart octave shift (fit to instrument range)',
                    'clamp': '    → Clamp remaining out-of-range notes',
                    'debass': '    → Remove bass duplicates of melody',
                    'simplify': '    → Simplify track (reduce filler / octave switches)',
                    'timing': '    → Adjust note timing (trim overlaps)',
                }
                for f in tr['fixes']:
                    lines.append(fix_labels.get(f, f'    → {f}'))
                if tr['is_melody']:
                    lines.append("    (melody track: only safe fixes applied)")

        # Melody-level fixes (not per-track)
        harm_sub = r.get('melody_harm_sub_candidates', 0)
        consol = r.get('melody_consolidation_candidates', 0)
        if harm_sub or consol:
            has_any_fix = True
            lines.append("\n─── Melody-wide fixes ───")
            if harm_sub:
                lines.append(f"  → Harmonic substitution: {harm_sub} notes (octave down + fifth)")
            if consol:
                lines.append(f"  → Octave consolidation: {consol} notes in rapid runs")

        if not has_any_fix:
            lines.append("\n✓ No fixes needed — song looks good!")
            self.fix_btn.setEnabled(False)

        self.text.setPlainText('\n'.join(lines))


# ── GUI ───────────────────────────────────────────────────────────────────────

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"Serenade Music Converter v{__version__}")
        self.settings = QSettings('Serenade', 'MIDIConverter')
        self.setMinimumSize(1100, 800)
        geo = self.settings.value('window_geometry')
        if geo:
            self.restoreGeometry(geo)
        else:
            self.resize(1200, 900)
        # App icon (check PyInstaller _MEIPASS first, then script dir)
        _base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        _icon_path = os.path.join(_base, 'serenade-icon.png')
        if os.path.isfile(_icon_path):
            self._app_icon = QIcon(_icon_path)
            self.setWindowIcon(self._app_icon)
        else:
            self._app_icon = None
        self.setAcceptDrops(True)
        self.midi_path = ""
        self.mid = None
        self.file_type = "midi"
        self.xml_root = None
        self._pr_notes = []  # PianoRollNote list
        self._current_instrument = GW2_INSTRUMENTS[0]  # default instrument
        self._pr_tracks = []  # track info dicts

        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setSpacing(6)
        main_layout.setContentsMargins(8, 8, 8, 8)

        # ── Menu bar ─────────────────────────────────────────────────
        menubar = self.menuBar()

        file_menu = menubar.addMenu("&File")
        new_menu = file_menu.addMenu("&New")
        for _inst_name, _inst_key, _inst_lo, _inst_hi in GW2_INSTRUMENTS:
            _oct = (_inst_hi - _inst_lo) // 12
            act = new_menu.addAction(f"{_inst_name}  ({_inst_key}, {_oct} oct)")
            _data = (_inst_name, _inst_key, _inst_lo, _inst_hi)
            act.triggered.connect(lambda checked, d=_data: self._new_song_instrument(d))
        _act = file_menu.addAction("&Import File...\tCtrl+O", self.browse_input)
        _act.setShortcut(QKeySequence("Ctrl+O"))
        self._recent_menu = file_menu.addMenu("&Recent Files")
        self._rebuild_recent_menu()
        file_menu.addSeparator()
        self._save_ahk_action = file_menu.addAction("Save &AHK\tCtrl+S", self.do_convert)
        self._save_ahk_action.setShortcut(QKeySequence("Ctrl+S"))
        self._save_ahk_action.setEnabled(False)
        file_menu.addAction("Export &MIDI...", self._export_midi)
        self._save_xml_action = file_menu.addAction("Save Music&XML", self.do_export_musicxml)
        self._save_xml_action.setEnabled(False)
        self._submit_action = file_menu.addAction("S&ubmit Song...", self.do_submit_song)
        self._submit_action.setEnabled(False)
        file_menu.addSeparator()
        file_menu.addAction("&Batch Convert...", self._batch_convert)
        file_menu.addSeparator()
        file_menu.addAction("&Quit", self.close)

        edit_menu = menubar.addMenu("&Edit")
        _act = edit_menu.addAction("&Undo\tCtrl+Z", lambda: self.piano_roll.undo())
        _act.setShortcut(QKeySequence("Ctrl+Z"))
        _act = edit_menu.addAction("&Redo\tCtrl+Y", lambda: self.piano_roll.redo())
        _act.setShortcut(QKeySequence("Ctrl+Y"))
        edit_menu.addSeparator()
        _act = edit_menu.addAction("&Copy\tCtrl+C", lambda: self.piano_roll.copySelected())
        _act.setShortcut(QKeySequence("Ctrl+C"))
        _act = edit_menu.addAction("&Paste\tCtrl+V", lambda: self.piano_roll.pasteClipboard())
        _act.setShortcut(QKeySequence("Ctrl+V"))
        edit_menu.addSeparator()
        _act = edit_menu.addAction("Select &All\tCtrl+A", lambda: [setattr(n, 'selected', True) for n in self._pr_notes if not n.deleted] or self.piano_roll.update())
        _act.setShortcut(QKeySequence("Ctrl+A"))
        _act = edit_menu.addAction("&Delete Selected\tDel", self._delete_selected_notes)
        _act.setShortcut(QKeySequence("Del"))

        view_menu = menubar.addMenu("&View")
        self._draw_action = view_menu.addAction("&Draw Mode\tCtrl+D")
        self._draw_action.setCheckable(True)
        self._draw_action.setShortcut(QKeySequence("Ctrl+D"))
        self._draw_action.toggled.connect(self._toggle_edit_mode)
        _act = view_menu.addAction("Zoom &In\tCtrl+=", lambda: self.piano_roll.zoomIn())
        _act.setShortcut(QKeySequence("Ctrl+="))
        _act = view_menu.addAction("Zoom &Out\tCtrl+-", lambda: self.piano_roll.zoomOut())
        _act.setShortcut(QKeySequence("Ctrl+-"))
        self._snap_action = view_menu.addAction("&Snap to Grid")
        self._snap_action.setCheckable(True)
        self._snap_action.setChecked(False)
        self._snap_action.toggled.connect(self._toggle_snap)

        tools_menu = menubar.addMenu("&Tools")
        tools_menu.aboutToShow.connect(self._update_tools_menu_state)
        self._tools_analyse_act = tools_menu.addAction("&Analyse Song...", self._analyse_song)
        tools_menu.addSeparator()
        self._tools_merge_act = tools_menu.addAction("&Merge Selected Tracks", self._merge_tracks)
        self._tools_split_act = tools_menu.addAction("S&plit Track by Pitch", self._split_track_by_pitch)
        self._tools_smart_oct_act = tools_menu.addAction("Smart &Octave Assignment", self._smart_octave)
        tools_menu.addSeparator()
        tools_menu.addAction("Remove &Duplicate Notes", self._remove_duplicates)
        tools_menu.addAction("Merge S&hort Gaps", self._merge_short_gaps)
        tools_menu.addAction("Edit Note &Velocity...", self._edit_velocity)

        help_menu = menubar.addMenu("&Help")
        help_menu.addAction("&User Guide", lambda: self._open_url("https://github.com/PieOrCake/serenade-converter/wiki"))
        help_menu.addAction("&Keyboard Shortcuts", self._show_shortcuts)
        help_menu.addAction("&About", self._show_about)

        # Hidden label for compatibility with code that sets input_label
        self.input_label = QLabel()
        self.input_label.hide()

        # ── Toolbar: metadata + instrument + mode ────────────────────
        toolbar2 = QHBoxLayout()
        toolbar2.setSpacing(8)

        toolbar2.addWidget(QLabel("Title:"))
        self.title_edit = QLineEdit()
        self.title_edit.setPlaceholderText("Song title")
        self.title_edit.setFixedWidth(200)
        toolbar2.addWidget(self.title_edit)

        toolbar2.addWidget(QLabel("Artist:"))
        self.author_edit = QLineEdit()
        self.author_edit.setPlaceholderText("Artist")
        self.author_edit.setFixedWidth(160)
        toolbar2.addWidget(self.author_edit)

        toolbar2.addWidget(QLabel("Instrument:"))
        self.instrument_combo = QComboBox()
        for _inst_name, _inst_key, _inst_lo, _inst_hi in GW2_INSTRUMENTS:
            _oct = (_inst_hi - _inst_lo) // 12
            self.instrument_combo.addItem(f"{_inst_name} ({_inst_key}, {_oct} oct)",
                                          (_inst_name, _inst_key, _inst_lo, _inst_hi))
        self.instrument_combo.setCurrentIndex(0)
        self.instrument_combo.currentIndexChanged.connect(self._on_instrument_changed)
        toolbar2.addWidget(self.instrument_combo)

        toolbar2.addStretch()

        # Update checker label (hidden until update found)
        self._update_label = QLabel()
        self._update_label.setTextFormat(Qt.TextFormat.RichText)
        self._update_label.setOpenExternalLinks(False)
        self._update_label.linkActivated.connect(self._open_url)
        self._update_label.hide()
        toolbar2.addWidget(self._update_label)

        self._mode_btn = QPushButton("⬚ Select")
        self._mode_btn.setCheckable(True)
        self._mode_btn.setMinimumHeight(30)
        self._mode_btn.setFixedWidth(120)
        self._mode_btn.setToolTip("Toggle between Select and Draw mode (Ctrl+D)")
        self._mode_btn.setStyleSheet(
            "QPushButton { background-color: #2a3a4a; border: 2px solid #4a6a8a; border-radius: 4px; font-weight: bold; padding: 2px 8px; }"
            "QPushButton:checked { background-color: #365; border-color: #6b6; color: #8f8; }")
        self._mode_btn.toggled.connect(self._on_draw_btn_toggled)
        toolbar2.addWidget(self._mode_btn)

        self._analyse_btn = QPushButton("🔍 Analyse")
        self._analyse_btn.setMinimumHeight(30)
        self._analyse_btn.setFixedWidth(120)
        self._analyse_btn.setToolTip("Analyse song for GW2 playback issues and auto-fix")
        self._analyse_btn.setStyleSheet(
            "QPushButton { background-color: #2a3a4a; border: 2px solid #4a6a8a; border-radius: 4px; font-weight: bold; padding: 2px 8px; }"
            "QPushButton:hover { background-color: #3a4a5a; }")
        self._analyse_btn.clicked.connect(self._analyse_song)
        toolbar2.addWidget(self._analyse_btn)

        # ☕ Buy me a coffee button with gentle pulsing burnt orange glow
        self._coffee_btn = QPushButton("☕ Buy me a coffee")
        self._coffee_btn.setMinimumHeight(30)
        self._coffee_btn.setFixedWidth(160)
        self._coffee_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self._coffee_btn.setToolTip("Support development on Ko-fi")
        self._coffee_btn.clicked.connect(lambda: self._open_url("https://ko-fi.com/pieorcake"))
        self._coffee_pulse_phase = 0.0
        self._coffee_pulse_count = 0  # pulses remaining in current burst
        self._coffee_idle_style = (
            "QPushButton { background-color: #2a3a4a; border: 2px solid #4a6a8a; "
            "border-radius: 4px; font-weight: bold; padding: 2px 8px; }"
            "QPushButton:hover { background-color: #3a4a5a; }")
        self._coffee_btn.setStyleSheet(self._coffee_idle_style)
        # Animation timer (runs only during pulse bursts)
        self._coffee_anim_timer = QTimer(self)
        self._coffee_anim_timer.timeout.connect(self._pulse_coffee_btn)
        # Schedule timer (fires to start a burst)
        self._coffee_schedule_timer = QTimer(self)
        self._coffee_schedule_timer.setSingleShot(True)
        self._coffee_schedule_timer.timeout.connect(self._start_coffee_burst)
        self._schedule_next_coffee_burst()
        toolbar2.addWidget(self._coffee_btn)

        # Hidden dummy widgets for enable/disable compatibility
        self.convert_btn = QPushButton(); self.convert_btn.hide()
        self.export_xml_btn = QPushButton(); self.export_xml_btn.hide()
        self.submit_btn = QPushButton(); self.submit_btn.hide()

        main_layout.addLayout(toolbar2)

        # ── Main content: left panel + piano roll ─────────────────────
        splitter = QSplitter(Qt.Orientation.Horizontal)

        # Left panel
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 4, 0)
        left_layout.setSpacing(6)

        # Track controls
        self.tracks_group = QGroupBox("Tracks")
        tracks_layout = QVBoxLayout(self.tracks_group)
        tracks_layout.setSpacing(2)

        track_btn_row = QHBoxLayout()
        self.select_all_btn = QPushButton("All")
        self.select_all_btn.setFixedWidth(50)
        self.select_all_btn.clicked.connect(self._select_all_tracks)
        self.select_none_btn = QPushButton("None")
        self.select_none_btn.setFixedWidth(50)
        self.select_none_btn.clicked.connect(self._select_no_tracks)
        track_btn_row.addWidget(self.select_all_btn)
        track_btn_row.addWidget(self.select_none_btn)
        track_btn_row.addStretch()
        tracks_layout.addLayout(track_btn_row)

        self.track_list = QListWidget()
        self.track_list.setMinimumHeight(120)
        self.track_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.track_list.customContextMenuRequested.connect(self._on_track_context_menu)
        tracks_layout.addWidget(self.track_list, 1)

        left_layout.addWidget(self.tracks_group, 1)

        # Conversion settings (collapsible)
        settings_group = QGroupBox("Conversion Settings")
        settings_group.setCheckable(True)
        settings_group.setChecked(False)
        settings_layout = QVBoxLayout(settings_group)
        settings_layout.setSpacing(4)

        settings_layout.addWidget(QLabel("Transpose:"))
        self.transpose_combo = QComboBox()
        self.transpose_combo.addItem("Auto (minimize octave changes)", None)
        for i in range(12):
            self.transpose_combo.addItem(f"+{i} semitones" if i else "None (keep original)", i)
        settings_layout.addWidget(self.transpose_combo)

        settings_layout.addWidget(QLabel("Chord window:"))
        self.chord_window_combo = QComboBox()
        self.chord_window_combo.addItem("Off (simultaneous only)", 5)
        self.chord_window_combo.addItem("Light (250ms)", 250)
        self.chord_window_combo.addItem("Medium (500ms)", 500)
        self.chord_window_combo.addItem("Heavy (850ms)", 850)
        settings_layout.addWidget(self.chord_window_combo)

        self.use_chords_cb = QCheckBox("Use GW2 chord mode")
        self.use_chords_cb.setToolTip("Substitute detected major/minor triads with\n"
                                       "GW2's built-in chord keypresses for better sound.")
        self.use_chords_cb.setChecked(False)
        settings_layout.addWidget(self.use_chords_cb)

        self.smooth_octaves_cb = QCheckBox("Smooth octave changes")
        self.smooth_octaves_cb.setToolTip("Flatten short octave excursions in fast passages.\n"
                                           "May improve simple melodies but can cause wrong\n"
                                           "pitches in complex arrangements.")
        self.smooth_octaves_cb.setChecked(False)
        settings_layout.addWidget(self.smooth_octaves_cb)

        left_layout.addWidget(settings_group)

        left_widget.setFixedWidth(260)
        splitter.addWidget(left_widget)

        # Piano Roll
        self.piano_roll = PianoRollWidget()
        splitter.addWidget(self.piano_roll)

        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        main_layout.addWidget(splitter, 1)

        self.piano_roll.selectionChanged.connect(self._update_play_btn_text)
        self.piano_roll.notesChanged.connect(self._on_notes_changed)
        self.piano_roll.pianoRollContextMenu.connect(self._on_piano_roll_context_menu)
        self.piano_roll.scrollChanged.connect(self._update_time_label)

        # ── Transport bar ─────────────────────────────────────────────
        transport = QHBoxLayout()
        transport.setSpacing(6)

        self.play_btn = QPushButton("▶ Play")
        self.play_btn.setMinimumHeight(30)
        self.play_btn.setFixedWidth(80)
        self.play_btn.setEnabled(False)
        self.play_btn.clicked.connect(self._playback_toggle)
        transport.addWidget(self.play_btn)

        self.play_here_btn = QPushButton("▶ Here")
        self.play_here_btn.setMinimumHeight(30)
        self.play_here_btn.setFixedWidth(80)
        self.play_here_btn.setEnabled(False)
        self.play_here_btn.setToolTip("Play from the current view position")
        self.play_here_btn.clicked.connect(self._playback_from_here)
        transport.addWidget(self.play_here_btn)

        self.stop_btn = QPushButton("■ Stop")
        self.stop_btn.setMinimumHeight(30)
        self.stop_btn.setFixedWidth(80)
        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self._playback_stop)
        transport.addWidget(self.stop_btn)

        # Hidden dummy for compatibility
        self.gw2_preview_cb = QCheckBox()
        self.gw2_preview_cb.hide()

        self._loop_cb = QCheckBox("Loop")
        self._loop_cb.setToolTip("Loop playback over the selected time range")
        transport.addWidget(self._loop_cb)

        self.playback_slider = QProgressBar()
        self.playback_slider.setRange(0, 1000)
        self.playback_slider.setValue(0)
        self.playback_slider.setTextVisible(False)
        self.playback_slider.setFixedHeight(16)
        self.playback_slider.hide()  # minimap handles progress display

        transport.addStretch(1)

        self._speed_spin = QSpinBox()
        self._speed_spin.setRange(10, 200)
        self._speed_spin.setSingleStep(10)
        self._speed_spin.setValue(100)
        self._speed_spin.setPrefix("Speed: ")
        self._speed_spin.setSuffix("%")
        self._speed_spin.setFixedWidth(110)
        self._speed_spin.setToolTip("Playback and export speed (10–200%)")
        self._speed_spin.valueChanged.connect(self._on_speed_changed)
        transport.addWidget(self._speed_spin)

        self.playback_time_label = QLabel("0:00 / 0:00")
        self.playback_time_label.setStyleSheet("color: #aaa; font-size: 22px;")
        transport.addWidget(self.playback_time_label)

        main_layout.addLayout(transport)

        # ── Stats bar ────────────────────────────────────────────────
        self._stats_label = QLabel("")
        self._stats_label.setStyleSheet("color: #888; font-size: 10px;")
        main_layout.addWidget(self._stats_label)

        # ── Minimap ──────────────────────────────────────────────────
        self._minimap = MinimapWidget(self)
        self._minimap.setFixedHeight(30)
        self._minimap.seekRequested.connect(self._on_minimap_seek)
        main_layout.insertWidget(main_layout.indexOf(splitter) + 1, self._minimap)

        # Playback state
        self._playback_sound = None
        self._playback_audio = None
        self._playback_total_ms = 0
        self._playback_start_time = 0
        self._playback_paused_at = 0
        self._playback_offset_ms = 0
        self._playback_playing = False
        self._playback_timer = QTimer()
        self._playback_timer.setInterval(30)
        self._playback_timer.timeout.connect(self._playback_tick)

        # ── Log ───────────────────────────────────────────────────────
        log_header = QHBoxLayout()
        log_header.addWidget(QLabel("Log"))
        log_header.addStretch()
        copy_log_btn = QPushButton("Copy")
        copy_log_btn.setFixedWidth(50)
        copy_log_btn.clicked.connect(lambda: QApplication.clipboard().setText(self.log_text.toPlainText()))
        log_header.addWidget(copy_log_btn)
        main_layout.addLayout(log_header)

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFont(QFont("monospace", 9))
        self.log_text.setMaximumHeight(80)
        self.log_text.setPlaceholderText("Log...")
        main_layout.addWidget(self.log_text)

        # Check for updates after UI is shown
        QTimer.singleShot(2000, self._check_for_updates)

    def log(self, msg):
        self.log_text.append(msg)

    def _populate_piano_roll(self):
        """Populate the piano roll from the currently loaded MIDI."""
        if not self.mid:
            return
        self.piano_roll._gw2_only = False
        # 3 octaves = 37 rows × 14px + ruler + scrollbar = 556px minimum
        self.piano_roll.setMinimumHeight(37 * self.piano_roll.NOTE_HEIGHT + self.piano_roll.RULER_HEIGHT + self.piano_roll.SCROLLBAR_SIZE)
        self.transpose_combo.setEnabled(True)
        self.chord_window_combo.setEnabled(True)
        self.tracks_group.setEnabled(True)
        # Reset GW2 range to current instrument selection
        data = self.instrument_combo.currentData()
        if data:
            self.piano_roll.GW2_PITCH_MIN = data[2]
            self.piano_roll.GW2_PITCH_MAX = data[3]
        notes, bpm = extract_notes_with_duration(self.mid)
        self._pr_notes = notes
        # Build track info
        track_indices = set(n.track for n in notes)
        tracks_info = get_track_info(self.midi_path)
        tracks_map = {idx: name for idx, name, count, rng, dup in tracks_info[0]}
        self._pr_tracks = []
        for tidx in sorted(track_indices):
            name = tracks_map.get(tidx, f"Track {tidx}")
            note_count = sum(1 for n in notes if n.track == tidx)
            if note_count == 0:
                continue
            color_idx = len(self._pr_tracks) % len(TRACK_COLORS)
            self._pr_tracks.append({
                'index': tidx,
                'name': name,
                'visible': True,
                'melody': False,
                'preserve': False,
                'simplify': False,
                'debass': False,
                'time_offset_ms': 0,
                'color': TRACK_COLORS[color_idx],
                'notes': note_count,
            })
        # Map note track indices to sequential indices for the piano roll
        track_remap = {}
        for i, t in enumerate(self._pr_tracks):
            track_remap[t['index']] = i
        for n in self._pr_notes:
            n.track = track_remap.get(n.track, 0)

        self.piano_roll.setNotes(self._pr_notes, self._pr_tracks, bpm)
        self._populate_track_list()
        self.play_btn.setEnabled(bool(self._pr_notes))
        self.play_here_btn.setEnabled(bool(self._pr_notes))

    def _populate_track_list(self):
        """Populate the track checkbox list from piano roll track info."""
        self.track_list.clear()
        for i, t in enumerate(self._pr_tracks):
            r, g, b = TRACK_COLORS[i % len(TRACK_COLORS)]
            melody_marker = "♪ " if t.get('melody', False) else ""
            preserve_marker = "♫ " if t.get('preserve', False) else ""
            simplify_marker = "✂ " if t.get('simplify', False) else ""
            debass_marker = "🔇 " if t.get('debass', False) else ""
            offset = t.get('time_offset_ms', 0)
            shift_marker = f"⏱{offset:+.0f}ms " if offset != 0 else ""
            label = f"{melody_marker}{preserve_marker}{simplify_marker}{debass_marker}{shift_marker}{t['name']} ({t['notes']})"
            item = QListWidgetItem(label)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Checked if t.get('visible', True) else Qt.CheckState.Unchecked)
            item.setData(Qt.ItemDataRole.UserRole, i)
            item.setForeground(QColor(r, g, b))
            self.track_list.addItem(item)
        # Connect itemChanged to toggle track visibility
        try:
            self.track_list.itemChanged.disconnect()
        except Exception:
            pass
        self.track_list.itemChanged.connect(self._on_track_toggled)

    def _on_track_toggled(self, item):
        idx = item.data(Qt.ItemDataRole.UserRole)
        visible = item.checkState() == Qt.CheckState.Checked
        if idx is not None and idx < len(self._pr_tracks):
            self._pr_tracks[idx]['visible'] = visible
            self.piano_roll.setTrackVisible(idx, visible)

    def _select_all_tracks(self):
        for i in range(self.track_list.count()):
            self.track_list.item(i).setCheckState(Qt.CheckState.Checked)

    def _select_no_tracks(self):
        for i in range(self.track_list.count()):
            self.track_list.item(i).setCheckState(Qt.CheckState.Unchecked)

    def _update_tools_menu_state(self):
        has_track = self.track_list.currentItem() is not None
        self._tools_merge_act.setEnabled(has_track)
        self._tools_split_act.setEnabled(has_track)
        self._tools_smart_oct_act.setEnabled(has_track)
        self._tools_analyse_act.setEnabled(bool(self._pr_notes))

    def _on_track_context_menu(self, pos):
        item = self.track_list.itemAt(pos)
        if item is None:
            return
        idx = item.data(Qt.ItemDataRole.UserRole)
        if idx is None:
            return
        menu = QMenu(self)
        menu.setStyleSheet(
            "QMenu::item { padding: 4px 12px 4px 20px; }"
            "QMenu::item:selected { background: #3a5a7a; color: #fff; }")
        name = self._pr_tracks[idx]['name'] if idx < len(self._pr_tracks) else f"Track {idx}"
        menu.addAction(f"Octave Up  ({name})", lambda: self._shift_track_octave(12))
        menu.addAction(f"Octave Down  ({name})", lambda: self._shift_track_octave(-12))
        clamp_sub = menu.addMenu("Clamp to Octave")
        clamp_sub.addAction("High (C5–B5)", lambda: self._clamp_to_octave(72))
        clamp_sub.addAction("Mid (C4–B4)", lambda: self._clamp_to_octave(60))
        clamp_sub.addAction("Low (C3–B3)", lambda: self._clamp_to_octave(48))
        menu.addSeparator()
        menu.addAction("Smart Octave Assignment", self._smart_octave)
        menu.addAction("Split by Pitch", self._split_track_by_pitch)
        menu.addAction("Merge Checked Tracks", self._merge_tracks)
        menu.addSeparator()
        is_melody = self._pr_tracks[idx].get('melody', False) if idx < len(self._pr_tracks) else False
        if is_melody:
            menu.addAction("♪ Clear Melody", lambda: self._set_melody_track(None))
        else:
            menu.addAction("♪ Set as Melody", lambda: self._set_melody_track(idx))
        is_preserved = self._pr_tracks[idx].get('preserve', False) if idx < len(self._pr_tracks) else False
        if is_preserved:
            menu.addAction("♫ Clear Preserve", lambda: self._set_track_preserve(idx, False))
        else:
            menu.addAction("♫ Preserve Track", lambda: self._set_track_preserve(idx, True))
        is_simplified = self._pr_tracks[idx].get('simplify', False) if idx < len(self._pr_tracks) else False
        if is_simplified:
            menu.addAction("✂ Clear Simplify", lambda: self._set_track_simplify(idx, False))
        else:
            menu.addAction("✂ Simplify (treble + bass)", lambda: self._set_track_simplify(idx, True))
        has_melody = any(t.get('melody', False) for t in self._pr_tracks)
        is_debassed = self._pr_tracks[idx].get('debass', False) if idx < len(self._pr_tracks) else False
        if is_debassed:
            menu.addAction("🔇 Clear Bass Removal", lambda: self._set_track_debass(idx, False))
        else:
            act = menu.addAction("🔇 Remove Bass Duplicates", lambda: self._set_track_debass(idx, True))
            if not has_melody:
                act.setEnabled(False)
                act.setToolTip("Set a melody track first")
        shift_sub = menu.addMenu("⏱ Time Shift")
        shift_sub.addAction("→ Shift Forward ⅛ Beat", lambda: self._shift_track_time(idx, +1))
        shift_sub.addAction("← Shift Back ⅛ Beat", lambda: self._shift_track_time(idx, -1))
        shift_sub.addSeparator()
        shift_sub.addAction("Reset Shift", lambda: self._shift_track_time(idx, 0))
        menu.addSeparator()
        menu.addAction("Select All Notes", lambda: self._select_track_notes(idx))
        menu.addAction("Delete Track Notes", lambda: self._delete_track_notes(idx))
        menu.exec(self.track_list.mapToGlobal(pos))

    def _select_track_notes(self, track_idx):
        for n in self._pr_notes:
            n.selected = (n.track == track_idx and not n.deleted)
        self.piano_roll.update()

    def _delete_track_notes(self, track_idx):
        self.piano_roll.pushUndo()
        count = 0
        for n in self._pr_notes:
            if n.track == track_idx and not n.deleted:
                n.deleted = True
                count += 1
        if count:
            self.piano_roll.notesChanged.emit()
            self.piano_roll.update()
            self.log(f"Deleted {count} notes from track {track_idx}")

    # ── Piano roll context menu (right-click) ────────────────────────
    def _get_target_notes(self):
        """Return (notes_list, label) for the current selection or all visible notes."""
        selected = [n for n in self._pr_notes if n.selected and not n.deleted]
        if selected:
            return selected, f"{len(selected)} selected"
        visible = set(i for i, t in enumerate(self._pr_tracks) if t.get('visible', True))
        all_vis = [n for n in self._pr_notes if not n.deleted and n.track in visible]
        return all_vis, f"all {len(all_vis)} notes"

    def _on_piano_roll_context_menu(self, global_pos):
        """Show context menu on right-click in the piano roll."""
        targets, label = self._get_target_notes()
        if not targets:
            return
        menu = QMenu(self)
        # Header showing scope
        header = menu.addAction(label.upper())
        header.setEnabled(False)
        hfont = header.font()
        hfont.setBold(True)
        hfont.setPointSize(hfont.pointSize() - 1)
        header.setFont(hfont)
        menu.setStyleSheet(
            "QMenu::item { padding: 4px 12px 4px 20px; }"
            "QMenu::item:selected { background: #3a5a7a; color: #fff; }"
            "QMenu::item:disabled { color: #8af; background: #2a3a4a; padding: 4px 12px; }")
        menu.addSeparator()
        menu.addAction("Octave Up", lambda: self._octave_shift_targets(12))
        menu.addAction("Octave Down", lambda: self._octave_shift_targets(-12))
        clamp_sub = menu.addMenu("Clamp to Octave")
        clamp_sub.addAction("High (C5–B5)", lambda: self._clamp_targets(72))
        clamp_sub.addAction("Mid (C4–B4)", lambda: self._clamp_targets(60))
        clamp_sub.addAction("Low (C3–B3)", lambda: self._clamp_targets(48))
        menu.addAction("Smart Octave Assignment", self._smart_octave_targets)
        menu.addSeparator()
        menu.addAction("✂ Simplify", self._simplify_targets)
        menu.addAction("✂ Unsimplify", self._unsimplify_targets)
        has_melody = any(t.get('melody', False) for t in self._pr_tracks)
        act = menu.addAction("🔇 Remove Bass Duplicates", self._debass_targets)
        if not has_melody:
            act.setEnabled(False)
        shift_sub = menu.addMenu("⏱ Time Shift")
        shift_sub.addAction("→ Shift Forward ⅛ Beat", lambda: self._time_shift_targets(+1))
        shift_sub.addAction("← Shift Back ⅛ Beat", lambda: self._time_shift_targets(-1))
        menu.addSeparator()
        menu.addAction("Delete", self._delete_targets)
        menu.exec(global_pos)

    def _octave_shift_targets(self, semitones):
        """Shift target notes by semitones (±12 = octave)."""
        targets, label = self._get_target_notes()
        if not targets:
            return
        self.piano_roll.pushUndo()
        for n in targets:
            n.pitch += semitones
        direction = "up" if semitones > 0 else "down"
        self.log(f"Shifted {label} {direction} one octave")
        self.piano_roll.notesChanged.emit()
        self.piano_roll.update()

    def _clamp_targets(self, base_pitch):
        """Clamp target notes into a single octave starting at base_pitch."""
        targets, label = self._get_target_notes()
        if not targets:
            return
        self.piano_roll.pushUndo()
        nn = ['C','C#','D','D#','E','F','F#','G','G#','A','A#','B']
        oct_name = f"{nn[base_pitch % 12]}{base_pitch // 12 - 1}–{nn[(base_pitch + 11) % 12]}{(base_pitch + 11) // 12 - 1}"
        count = 0
        for n in targets:
            new_pitch = base_pitch + (n.pitch % 12)
            if new_pitch != n.pitch:
                n.pitch = new_pitch
                count += 1
        self.piano_roll.notesChanged.emit()
        self.piano_roll.update()
        self.log(f"Clamped {count} of {label} to {oct_name}")

    def _smart_octave_targets(self):
        """Assign target notes to the octave that minimizes out-of-range notes."""
        targets, label = self._get_target_notes()
        if not targets:
            return
        inst_data = self.instrument_combo.currentData()
        if not inst_data:
            return
        lo, hi = inst_data[2], inst_data[3]
        self.piano_roll.pushUndo()
        # Group targets by track
        by_track = {}
        for n in targets:
            by_track.setdefault(n.track, []).append(n)
        total_shifted = 0
        for tidx, tnotes in by_track.items():
            best_shift = 0
            best_oor = len(tnotes)
            for shift in range(-48, 49, 12):
                oor = sum(1 for n in tnotes if (n.pitch + shift) < lo or (n.pitch + shift) > hi)
                if oor < best_oor:
                    best_oor = oor
                    best_shift = shift
            if best_shift != 0:
                for n in tnotes:
                    n.pitch += best_shift
                total_shifted += len(tnotes)
        self.piano_roll.notesChanged.emit()
        self.piano_roll.update()
        self.log(f"Smart octave: adjusted {total_shifted} of {label}")

    def _simplify_targets(self):
        """Manually mark target notes as simplified."""
        targets, label = self._get_target_notes()
        if not targets:
            return
        self.piano_roll.pushUndo()
        count = 0
        for n in targets:
            if n.simplified_manual is not True:
                n.simplified_manual = True
                count += 1
        self.piano_roll.updateSimplifiedNotes()
        self.piano_roll.update()
        self.log(f"✂ Simplified {count} of {label}")

    def _unsimplify_targets(self):
        """Clear simplification on target notes."""
        targets, label = self._get_target_notes()
        if not targets:
            return
        self.piano_roll.pushUndo()
        count = 0
        for n in targets:
            if n.simplified_manual is not None:
                n.simplified_manual = None
                count += 1
        self.piano_roll.updateSimplifiedNotes()
        self.piano_roll.update()
        self.log(f"✂ Cleared simplify overrides on {count} of {label}")

    def _debass_targets(self):
        """Remove lower-octave duplicates from targets relative to melody track."""
        targets, label = self._get_target_notes()
        if not targets:
            return
        melody_tracks = set(i for i, t in enumerate(self._pr_tracks) if t.get('melody', False))
        if not melody_tracks:
            self.log("Set a melody track first.")
            return
        visible = set(i for i, t in enumerate(self._pr_tracks) if t.get('visible', True))
        offsets = self.piano_roll._track_offsets()
        CHORD_WIN = 5
        # Collect melody notes
        melody_notes = [n for n in self._pr_notes
                        if not n.deleted and not n.simplified
                        and n.track in melody_tracks and n.track in visible]
        melody_notes.sort(key=lambda n: n.start_ms + offsets.get(n.track, 0))
        if not melody_notes:
            self.log("No melody notes found.")
            return
        # Build time buckets
        buckets = []
        i = 0
        while i < len(melody_notes):
            bucket_start = melody_notes[i].start_ms + offsets.get(melody_notes[i].track, 0)
            pitches = set()
            j = i
            while j < len(melody_notes) and (melody_notes[j].start_ms + offsets.get(melody_notes[j].track, 0)) - bucket_start <= CHORD_WIN:
                pitches.add((melody_notes[j].pitch % 12, melody_notes[j].pitch))
                j += 1
            buckets.append((bucket_start, pitches))
            i = j
        # Check each target note (skip melody notes themselves)
        self.piano_roll.pushUndo()
        target_set = set(id(n) for n in targets)
        non_melody = [n for n in targets if n.track not in melody_tracks]
        non_melody.sort(key=lambda n: n.start_ms + offsets.get(n.track, 0))
        count = 0
        bi = 0
        for n in non_melody:
            n_ms = n.start_ms + offsets.get(n.track, 0)
            while bi > 0 and bi < len(buckets) and buckets[bi][0] > n_ms + CHORD_WIN:
                bi -= 1
            while bi < len(buckets) and buckets[bi][0] < n_ms - CHORD_WIN:
                bi += 1
            for bk in range(bi, len(buckets)):
                bstart, bpitches = buckets[bk]
                if bstart > n_ms + CHORD_WIN:
                    break
                if abs(bstart - n_ms) <= CHORD_WIN:
                    pc = n.pitch % 12
                    for mel_pc, mel_pitch in bpitches:
                        if pc == mel_pc and n.pitch < mel_pitch:
                            n.simplified_manual = True
                            count += 1
                            break
                if n.simplified_manual is True:
                    break
        self.piano_roll.updateSimplifiedNotes()
        self.piano_roll.update()
        self.log(f"🔇 Bass removal: {count} lower-octave duplicates hidden in {label}")

    def _time_shift_targets(self, direction):
        """Destructive time shift on target notes by ⅛ beat."""
        targets, label = self._get_target_notes()
        if not targets:
            return
        bpm = self.piano_roll._bpm or 120
        eighth_ms = (60000.0 / bpm) / 2
        self.piano_roll.pushUndo()
        for n in targets:
            n.start_ms = max(0, n.start_ms + direction * eighth_ms)
        direction_label = "forward" if direction > 0 else "back"
        self.piano_roll.notesChanged.emit()
        self.piano_roll.update()
        self.log(f"⏱ Shifted {label} {direction_label} ⅛ beat ({direction * eighth_ms:+.0f} ms)")

    def _delete_targets(self):
        """Delete target notes."""
        targets, label = self._get_target_notes()
        if not targets:
            return
        self.piano_roll.pushUndo()
        for n in targets:
            n.deleted = True
            n.selected = False
        self.piano_roll.notesChanged.emit()
        self.piano_roll.update()
        self.log(f"Deleted {label}")

    def _set_track_simplify(self, track_idx, enabled):
        """Toggle simplify on a track. Clears manual overrides when disabling."""
        if track_idx < len(self._pr_tracks):
            self._pr_tracks[track_idx]['simplify'] = enabled
            if not enabled:
                # Clear manual overrides for this track
                for n in self._pr_notes:
                    if n.track == track_idx:
                        n.simplified_manual = None
            self._populate_track_list()
            self.piano_roll.updateSimplifiedNotes()
            self.piano_roll.update()
            name = self._pr_tracks[track_idx]['name']
            count = sum(1 for n in self._pr_notes if n.simplified and n.track == track_idx)
            if enabled:
                self.log(f"✂ Simplify {name}: {count} middle notes hidden")
            else:
                self.log(f"✂ Simplify cleared: {name}")

    def _set_melody_track(self, track_idx):
        """Set the given track as melody (or clear all if track_idx is None)."""
        for i, t in enumerate(self._pr_tracks):
            t['melody'] = (i == track_idx)
        self._populate_track_list()
        self.piano_roll.updateSimplifiedNotes()
        self.piano_roll.update()
        if track_idx is not None and track_idx < len(self._pr_tracks):
            self.log(f"♪ Melody track: {self._pr_tracks[track_idx]['name']}")
        else:
            self.log("♪ Melody track cleared")

    def _set_track_preserve(self, track_idx, enabled):
        """Toggle preserve flag on a track (exempt from debass/simplify/muting)."""
        if track_idx >= len(self._pr_tracks):
            return
        self._pr_tracks[track_idx]['preserve'] = enabled
        self._populate_track_list()
        self.piano_roll.update()
        name = self._pr_tracks[track_idx]['name']
        if enabled:
            self.log(f"♫ Preserve: {name}")
        else:
            self.log(f"♫ Preserve cleared: {name}")

    def _set_track_debass(self, track_idx, enabled):
        """Toggle bass duplicate removal on a track (relative to the melody track)."""
        if track_idx >= len(self._pr_tracks):
            return
        self._pr_tracks[track_idx]['debass'] = enabled
        if not enabled:
            for n in self._pr_notes:
                if n.track == track_idx:
                    n.simplified_manual = None
        self._populate_track_list()
        self.piano_roll.updateSimplifiedNotes()
        self.piano_roll.update()
        name = self._pr_tracks[track_idx]['name']
        if enabled:
            count = sum(1 for n in self._pr_notes if n.simplified and n.track == track_idx)
            self.log(f"🔇 Bass removal {name}: {count} lower-octave duplicates hidden")
        else:
            self.log(f"🔇 Bass removal cleared: {name}")

    def _shift_track_time(self, track_idx, direction):
        """Shift a track's time offset by ⅛ beat. direction: +1 forward, -1 back, 0 reset."""
        if track_idx >= len(self._pr_tracks):
            return
        bpm = self.piano_roll._bpm or 120
        eighth_ms = (60000.0 / bpm) / 2  # ⅛ beat in ms
        t = self._pr_tracks[track_idx]
        if direction == 0:
            t['time_offset_ms'] = 0
        else:
            t['time_offset_ms'] = t.get('time_offset_ms', 0) + direction * eighth_ms
        self._populate_track_list()
        self.piano_roll.updateSimplifiedNotes()
        self.piano_roll.update()
        self._minimap.update()
        offset = t['time_offset_ms']
        if offset == 0:
            self.log(f"⏱ Time shift reset: {t['name']}")
        else:
            beats = offset / eighth_ms
            self.log(f"⏱ Time shift {t['name']}: {beats:+.0f}/8 beat ({offset:+.0f} ms)")

    def _clamp_to_octave(self, base_pitch):
        """Clamp selected notes (or selected track's notes) into a single octave starting at base_pitch."""
        # Determine which notes to clamp: selected notes first, else selected track
        targets = [n for n in self._pr_notes if n.selected and not n.deleted]
        source = "selected notes"
        if not targets:
            item = self.track_list.currentItem()
            if item is None:
                self.log("Select notes or a track first.")
                return
            idx = item.data(Qt.ItemDataRole.UserRole)
            if idx is None:
                return
            targets = [n for n in self._pr_notes if n.track == idx and not n.deleted]
            source = f"track '{self._pr_tracks[idx]['name']}'"
        if not targets:
            self.log("No notes to clamp.")
            return
        self.piano_roll.pushUndo()
        nn = ['C','C#','D','D#','E','F','F#','G','G#','A','A#','B']
        oct_name = f"{nn[base_pitch % 12]}{base_pitch // 12 - 1}–{nn[(base_pitch + 11) % 12]}{(base_pitch + 11) // 12 - 1}"
        count = 0
        for n in targets:
            new_pitch = base_pitch + (n.pitch % 12)
            if new_pitch != n.pitch:
                n.pitch = new_pitch
                count += 1
        self.piano_roll.update()
        self.piano_roll.notesChanged.emit()
        self.log(f"Clamped {count} of {len(targets)} {source} to {oct_name}")

    def _shift_track_octave(self, semitones):
        """Shift all notes in the selected track by the given number of semitones."""
        item = self.track_list.currentItem()
        if item is None:
            self.log("Select a track first.")
            return
        idx = item.data(Qt.ItemDataRole.UserRole)
        if idx is None:
            return
        self.piano_roll.pushUndo()
        count = 0
        for note in self._pr_notes:
            if note.track == idx and not note.deleted:
                note.pitch += semitones
                count += 1
        if count:
            direction = "up" if semitones > 0 else "down"
            self.log(f"Shifted {count} notes in '{self._pr_tracks[idx]['name']}' {direction} one octave")
            self.piano_roll.update()

    def _new_song_instrument(self, inst_data):
        """Create a new blank song for a specific GW2 instrument."""
        inst_name, inst_key, midi_lo, midi_hi = inst_data
        self.midi_path = ""
        self.mid = None
        self.xml_root = None
        self.file_type = "composed"
        self._current_instrument = inst_data
        self.piano_roll._undo_stack.clear()
        self.piano_roll._redo_stack.clear()
        self._pr_notes = []
        self._pr_tracks = [{
            'index': 0,
            'name': inst_name,
            'visible': True,
            'melody': True,
            'simplify': False,
            'color': TRACK_COLORS[0],
            'notes': 0,
        }]
        bpm = 120
        self.piano_roll.setInstrumentRange(midi_lo, midi_hi)
        self.piano_roll.setNotes(self._pr_notes, self._pr_tracks, bpm)
        self.transpose_combo.setEnabled(False)
        self.chord_window_combo.setEnabled(False)
        self.tracks_group.setEnabled(False)
        self.piano_roll._total_ms = 32 * 4 * (60000.0 / bpm)  # 32 bars
        self.piano_roll._scroll_x = 0
        self.piano_roll._scroll_y = 0
        self.piano_roll._update_scrollbars()
        self.piano_roll.update()

        self._populate_track_list()
        self.play_btn.setEnabled(True)
        self.play_here_btn.setEnabled(True)
        self._enable_export(ahk=True, xml=False, submit=True)

        nn = ['C','C#','D','D#','E','F','F#','G','G#','A','A#','B']
        lo_name = f"{nn[midi_lo % 12]}{midi_lo // 12 - 1}"
        hi_name = f"{nn[midi_hi % 12]}{midi_hi // 12 - 1}"
        octaves = (midi_hi - midi_lo) // 12

        self.setWindowTitle(f"Serenade Music Converter v{__version__} — New: {inst_name}")
        self.title_edit.clear()
        self.author_edit.clear()
        self.log_text.clear()
        self.log(f"New {inst_name} song created (120 BPM, 32 bars)")
        self.log(f"Key: {inst_key} | Range: {lo_name}-{hi_name} ({octaves} octaves)")
        self.log("Switch to Draw mode to start placing notes.")
        self.log("Right-click to delete notes. Ctrl+Z to undo.")

        # Set instrument combo to match
        for i in range(self.instrument_combo.count()):
            d = self.instrument_combo.itemData(i)
            if d and d[0] == inst_name:
                self.instrument_combo.blockSignals(True)
                self.instrument_combo.setCurrentIndex(i)
                self.instrument_combo.blockSignals(False)
                break

        # Auto-switch to draw mode
        self._mode_btn.setChecked(True)

    def _show_about(self):
        from PyQt6.QtWidgets import QDialog, QDialogButtonBox
        import webbrowser
        dlg = QDialog(self)
        dlg.setWindowTitle("About Serenade Music Converter")
        layout = QVBoxLayout(dlg)
        if self._app_icon:
            icon_label = QLabel()
            icon_label.setPixmap(self._app_icon.pixmap(64, 64))
            icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            layout.addWidget(icon_label)
        text_label = QLabel(
            f"<h3>Serenade Music Converter v{__version__}</h3>"
            "<p>Convert MIDI, MusicXML, and AHK files into GW2-compatible<br>"
            "instrument scripts with a full-featured piano roll editor.</p>"
            "<p><b>Features:</b> Multi-track editing, per-track simplify,<br>"
            "melody priority, auto-transpose, GW2 chord mode,<br>"
            "octave smoothing, MIDI playback preview, drag &amp; drop.</p>"
            "<p>Licensed under the <b>GNU General Public License v3.0</b> (GPL-3.0).<br>"
            "This is free software; you are free to change and redistribute it.<br>"
            "There is NO WARRANTY, to the extent permitted by law.</p>"
            "<p><a href='https://www.gnu.org/licenses/gpl-3.0.html'>https://www.gnu.org/licenses/gpl-3.0.html</a></p>"
            "<p><a href='https://pie.rocks.cc/'>Homepage: pie.rocks.cc</a></p>"
            "<p><a href='https://ko-fi.com/pieorcake'>☕ Buy me a coffee!</a></p>")
        text_label.setTextFormat(Qt.TextFormat.RichText)
        text_label.setWordWrap(True)
        text_label.setOpenExternalLinks(False)
        def _open_link(url):
            import subprocess, platform
            try:
                if platform.system() == 'Darwin':
                    subprocess.Popen(['open', url])
                elif platform.system() == 'Windows':
                    os.startfile(url)
                else:
                    # Linux: clear AppImage's LD_LIBRARY_PATH to avoid lib conflicts
                    env = dict(os.environ)
                    for key in ('LD_LIBRARY_PATH', 'LD_PRELOAD'):
                        env.pop(key, None)
                    subprocess.Popen(['xdg-open', url], env=env)
            except Exception as e:
                import webbrowser
                webbrowser.open(url)
        text_label.linkActivated.connect(_open_link)
        layout.addWidget(text_label)
        btn_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
        btn_box.accepted.connect(dlg.accept)
        layout.addWidget(btn_box)
        dlg.exec()

    def _on_instrument_changed(self, index):
        """Update GW2 range boundaries when instrument changes."""
        data = self.instrument_combo.currentData()
        if not data:
            return
        inst_name, inst_key, midi_lo, midi_hi = data
        self._current_instrument = data
        if self.file_type in ("composed", "ahk"):
            self.piano_roll.setInstrumentRange(midi_lo, midi_hi)
        else:
            self.piano_roll.GW2_PITCH_MIN = midi_lo
            self.piano_roll.GW2_PITCH_MAX = midi_hi
            self.piano_roll.update()

    def _on_notes_changed(self):
        """Called when notes are added/deleted/modified in the piano roll."""
        # Update track note counts
        for t in self._pr_tracks:
            tidx_seq = self._pr_tracks.index(t)
            t['notes'] = sum(1 for n in self._pr_notes if not n.deleted and n.track == tidx_seq)
        self._populate_track_list()
        self._update_stats()

    def _enable_export(self, ahk=True, xml=True, submit=True):
        """Enable/disable export buttons and corresponding menu actions."""
        self.convert_btn.setEnabled(ahk)
        self.export_xml_btn.setEnabled(xml)
        self.submit_btn.setEnabled(submit)
        self._save_ahk_action.setEnabled(ahk)
        self._save_xml_action.setEnabled(xml)
        self._submit_action.setEnabled(submit)

    def _delete_selected_notes(self):
        """Delete all selected notes in the piano roll."""
        has_selected = any(n.selected for n in self._pr_notes)
        if has_selected:
            self.piano_roll.pushUndo()
            for n in self._pr_notes:
                if n.selected:
                    n.deleted = True
                    n.selected = False
            self.piano_roll.notesChanged.emit()
            self.piano_roll.update()

    def _show_shortcuts(self):
        """Show keyboard shortcuts dialog."""
        QMessageBox.information(self, "Keyboard Shortcuts",
            "FILE\n"
            "Import file: Ctrl+O\n"
            "Save AHK: Ctrl+S\n"
            "Drag file onto window: Load file\n\n"
            "EDIT\n"
            "Undo: Ctrl+Z\n"
            "Redo: Ctrl+Y / Ctrl+Shift+Z\n"
            "Copy: Ctrl+C\n"
            "Paste: Ctrl+V\n"
            "Select all: Ctrl+A\n"
            "Deselect all: Escape\n"
            "Delete selected: Del\n\n"
            "VIEW\n"
            "Draw mode: Ctrl+D\n"
            "Zoom in: Ctrl+= / Ctrl+wheel up\n"
            "Zoom out: Ctrl+- / Ctrl+wheel down\n"
            "Scroll: Mouse wheel\n"
            "Horizontal scroll: Shift + Mouse wheel\n"
            "Go to start: Home\n"
            "Go to end: End\n\n"
            "PLAYBACK\n"
            "Play / Pause: Space\n"
            "Click minimap: Scroll to position\n\n"
            "PIANO ROLL\n"
            "Select notes: Click / Drag box\n"
            "Multi-select: Ctrl+Click\n"
            "Toggle simplified: Ctrl+Shift+Click (on simplified tracks)\n"
            "Select time range: Click+drag on ruler\n"
            "Right-click note: Delete note\n"
            "Right-click ruler: Trim menu\n"
            "Drag note edge: Resize note\n\n"
            "TRACK LIST (right-click)\n"
            "Set as Melody / Clear Melody\n"
            "Simplify (treble + bass) / Clear Simplify\n"
            "Octave Up / Down, Clamp to Octave\n"
            "Smart Octave Assignment\n"
            "Split by Pitch, Merge Checked Tracks\n"
            "Select All Notes, Delete Track Notes")

    def _schedule_next_coffee_burst(self):
        """Schedule the next pulse burst at a random interval (5-15 minutes)."""
        import random
        delay_ms = random.randint(5 * 60 * 1000, 15 * 60 * 1000)
        self._coffee_schedule_timer.start(delay_ms)

    def _start_coffee_burst(self):
        """Start a 3-pulse burst animation."""
        self._coffee_pulse_phase = 0.0
        self._coffee_pulse_count = 3
        self._coffee_anim_timer.start(50)

    def _pulse_coffee_btn(self):
        """Animate one tick of the pulse burst."""
        import math
        self._coffee_pulse_phase += 0.12
        t = (math.sin(self._coffee_pulse_phase) + 1.0) / 2.0  # 0..1
        # Check if we completed a full sine cycle (crossed zero going positive)
        if self._coffee_pulse_phase >= self._coffee_pulse_count * 2 * math.pi:
            self._coffee_anim_timer.stop()
            self._coffee_btn.setStyleSheet(self._coffee_idle_style)
            self._schedule_next_coffee_burst()
            return
        # Normal: bg #2a3a4a (42,58,74), border #4a6a8a (74,106,138), text #c8c8c8 (200,200,200)
        # Target: bg #5a3018 (90,48,24), border #cc7030 (204,112,48), text #e8a050 (232,160,80)
        def _lerp(a, b): return int(a + t * (b - a))
        bgr, bgg, bgb = _lerp(42, 90), _lerp(58, 48), _lerp(74, 24)
        bdr, bdg, bdb = _lerp(74, 204), _lerp(106, 112), _lerp(138, 48)
        txr, txg, txb = _lerp(200, 232), _lerp(200, 160), _lerp(200, 80)
        self._coffee_btn.setStyleSheet(
            f"QPushButton {{ background-color: rgb({bgr},{bgg},{bgb}); "
            f"border: 2px solid rgb({bdr},{bdg},{bdb}); border-radius: 4px; "
            f"font-weight: bold; padding: 2px 8px; color: rgb({txr},{txg},{txb}); }}"
            f"QPushButton:hover {{ background-color: rgb({min(255,bgr+30)},{min(255,bgg+20)},{min(255,bgb+10)}); }}")

    def _open_url(self, url):
        """Open a URL in the system browser, handling AppImage environment."""
        import subprocess, platform
        try:
            if platform.system() == 'Darwin':
                subprocess.Popen(['open', url])
            elif platform.system() == 'Windows':
                os.startfile(url)
            else:
                env = dict(os.environ)
                for key in ('LD_LIBRARY_PATH', 'LD_PRELOAD'):
                    env.pop(key, None)
                subprocess.Popen(['xdg-open', url], env=env)
        except Exception:
            import webbrowser
            webbrowser.open(url)

    def _check_for_updates(self):
        """Check GitHub for a newer converter release in a background thread."""
        import threading
        self._latest_version = None
        def _fetch():
            try:
                url = "https://api.github.com/repos/PieOrCake/serenade-converter/releases/latest"
                req = urllib.request.Request(url, headers={"User-Agent": "Serenade-Converter"})
                with urllib.request.urlopen(req, timeout=10) as resp:
                    data = json.loads(resp.read().decode())
                    self._latest_version = data.get("tag_name", "").lstrip("v")
            except Exception:
                pass
        t = threading.Thread(target=_fetch, daemon=True)
        t.start()
        # Poll for result without blocking the UI
        def _poll():
            if t.is_alive():
                QTimer.singleShot(500, _poll)
                return
            self._on_update_result()
        QTimer.singleShot(500, _poll)

    def _on_update_result(self):
        """Called when the background version check finishes."""
        if not self._latest_version:
            return
        current = tuple(int(x) for x in __version__.split('.'))
        try:
            latest = tuple(int(x) for x in self._latest_version.split('.'))
        except ValueError:
            return
        if latest > current:
            self._update_label.setText(
                f"<a href='https://pie.rocks.cc/projects/serenade-converter/' "
                f"style='color: #e8a050; text-decoration: none; font-weight: bold;'>"
                f"⬆ Update available: v{self._latest_version}</a>")
            self._update_label.show()

    def _on_draw_btn_toggled(self, checked):
        """Sync draw mode between toolbar button and menu action."""
        self._draw_action.blockSignals(True)
        self._draw_action.setChecked(checked)
        self._draw_action.blockSignals(False)
        self._toggle_edit_mode(checked)

    def _toggle_edit_mode(self, checked):
        self._mode_btn.blockSignals(True)
        self._mode_btn.setChecked(checked)
        self._mode_btn.blockSignals(False)
        if checked:
            self._mode_btn.setText("✏ Draw")
            self.piano_roll.setEditMode('draw')
        else:
            self._mode_btn.setText("⬚ Select")
            self.piano_roll.setEditMode('select')

    def browse_input(self):
        last_input_dir = self.settings.value('last_input_dir', '')
        path, _ = QFileDialog.getOpenFileName(
            self, "Select Music File", last_input_dir,
            "Music Files (*.mid *.midi *.musicxml *.mxl *.ahk);;MIDI Files (*.mid *.midi);;MusicXML Files (*.musicxml *.mxl);;AHK Scripts (*.ahk);;All Files (*)")
        if not path:
            return
        self._load_file(path)

    def _load_file(self, path):
        """Load a music file by path (used by browse, drag-drop, and recent files)."""
        if not os.path.isfile(path):
            self.log(f"File not found: {path}")
            return
        self.settings.setValue('last_input_dir', os.path.dirname(path))
        self._add_recent_file(path)
        ext = os.path.splitext(path)[1].lower()

        self.midi_path = path
        self.setWindowTitle(f"Serenade Music Converter v{__version__} — {os.path.basename(path)}")
        self.piano_roll._undo_stack.clear()
        self.piano_roll._redo_stack.clear()

        if ext == '.ahk':
            self.file_type = "ahk"
            self.mid = None
            self.xml_root = None
            try:
                notes, bpm, title, author = parse_ahk_to_notes(path)
                self._pr_notes = notes
                self._pr_tracks = [{
                    'index': 0,
                    'name': 'AHK Script',
                    'visible': True,
                    'melody': True,
                    'simplify': False,
                    'color': TRACK_COLORS[0],
                    'notes': len(notes),
                }]
                self.piano_roll.setInstrumentRange(48, 84)
                self.piano_roll.setNotes(self._pr_notes, self._pr_tracks, bpm)
                self.transpose_combo.setEnabled(False)
                self.chord_window_combo.setEnabled(False)
                self.tracks_group.setEnabled(False)
                self._populate_track_list()
                self.play_btn.setEnabled(bool(self._pr_notes))
                self.play_here_btn.setEnabled(bool(self._pr_notes))
                if title:
                    self.title_edit.setText(title)
                if author:
                    self.author_edit.setText(author)
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to parse AHK: {e}")
                return
        elif ext in ('.musicxml', '.mxl', '.xml'):
            self.file_type = "musicxml"
            try:
                self.xml_root = parse_musicxml_file(path)
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to parse MusicXML: {e}")
                return
            # For MusicXML, also try to load as MIDI via music21 for piano roll
            self.mid = None
            try:
                from music21 import converter as m21converter
                import tempfile
                score = m21converter.parse(path)
                tmp = tempfile.NamedTemporaryFile(suffix='.mid', delete=False)
                score.write('midi', fp=tmp.name)
                tmp.close()
                self.mid = mido.MidiFile(tmp.name)
                self._populate_piano_roll()
            except Exception:
                pass  # Piano roll won't show for MusicXML without music21
        else:
            self.file_type = "midi"
            self.xml_root = None
            try:
                self.mid = mido.MidiFile(path)
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to read MIDI: {e}")
                return
            try:
                self._populate_piano_roll()
            except Exception as e:
                self.log(f"ERROR loading MIDI: {e}")
                import traceback; traceback.print_exc()
                return

        # Auto-fill metadata
        basename = os.path.splitext(os.path.basename(path))[0]
        basename = re.sub(r'[\u2010-\u2015\u2212]', '-', basename)

        if self.file_type == "ahk":
            pass  # title/author already set from AHK metadata
        elif self.file_type == "musicxml" and self.xml_root is not None:
            meta = extract_musicxml_metadata(self.xml_root)
            title, artist = meta.get('title', ''), meta.get('artist', '')
            if title:
                self.title_edit.setText(title)
                self.author_edit.setText(artist)
            else:
                self.title_edit.setText(basename)
                self.author_edit.clear()
        else:
            suggestions = generate_metadata_suggestions(path, self.mid)
            if suggestions:
                self.title_edit.setText(suggestions[0][0])
                self.author_edit.setText(suggestions[0][1])
            else:
                self.title_edit.setText(basename)
                self.author_edit.clear()

        # Auto-fill output path
        last_output_dir = self.settings.value('last_output_dir', '')
        if last_output_dir:
            pass  # output path set at save time
        else:
            pass  # output path set at save time

        self._enable_export(ahk=True, xml=(self.file_type == "midi"), submit=True)
        self.log_text.clear()
        self.log(f"Loaded: {os.path.basename(path)}")
        if self._pr_tracks:
            self.log(f"Tracks: {len(self._pr_tracks)}, Notes: {len([n for n in self._pr_notes if not n.deleted])}")
        self._update_stats()

    def _flash_button(self, btn, text, color):
        original_text = btn.text()
        original_style = btn.styleSheet()
        btn.setText(text)
        btn.setStyleSheet(f"background-color: {color};")
        QTimer.singleShot(1500, lambda: (btn.setText(original_text), btn.setStyleSheet(original_style)))

    def do_submit_song(self):
        """Open a dialog explaining the submission process, then open the GitHub issue template."""
        title = self.title_edit.text().strip() or ""
        author = self.author_edit.text().strip() or ""
        inst_data = self.instrument_combo.currentData()
        instrument = inst_data[0] if inst_data else "Ornate Grand Piano"

        dlg = QMessageBox(self)
        dlg.setWindowTitle("Submit Song")
        dlg.setIcon(QMessageBox.Icon.Information)
        dlg.setText(
            "This will open a GitHub page where you can submit your song "
            "to the Serenade music library.\n\n"
            "What happens next:\n"
            "1. Your browser will open a pre-filled song submission form\n"
            "2. Drag and drop your saved .ahk file into the file field\n"
            "3. Review the details and click 'Submit new issue'\n\n"
            "A GitHub account is required. Your submission will be reviewed "
            "before being added to the library."
        )
        continue_btn = dlg.addButton("Continue", QMessageBox.ButtonRole.AcceptRole)
        dlg.addButton("Cancel", QMessageBox.ButtonRole.RejectRole)
        dlg.exec()

        if dlg.clickedButton() != continue_btn:
            return

        params = {
            "template": "song-submission.yml",
            "title": f"[Song] {title}" if title else "[Song] ",
        }
        if title:
            params["song-title"] = title
        if author:
            params["artist"] = author
        if instrument:
            params["instrument"] = instrument

        url = "https://github.com/PieOrCake/serenade/issues/new?" + urllib.parse.urlencode(params)
        webbrowser.open(url)
        self.log("Opened song submission page in browser.")

    def do_export_musicxml(self):
        """Convert loaded MIDI to MusicXML using music21."""
        if not self.midi_path:
            QMessageBox.warning(self, "Warning", "No MIDI file loaded.")
            return

        base = os.path.splitext(self.midi_path)[0]
        default_path = base + '.musicxml'
        last_output_dir = self.settings.value('last_output_dir', '')
        if last_output_dir:
            default_path = os.path.join(last_output_dir,
                os.path.splitext(os.path.basename(self.midi_path))[0] + '.musicxml')

        path, _ = QFileDialog.getSaveFileName(
            self, "Save MusicXML File", default_path,
            "MusicXML Files (*.musicxml);;All Files (*)")
        if not path:
            return

        self.log_text.clear()
        self.log("Converting MIDI to MusicXML...")
        self.export_xml_btn.setEnabled(False)
        QApplication.processEvents()

        try:
            from music21 import converter as m21converter
            from music21 import stream, note as m21note, meter
            score = m21converter.parse(self.midi_path)

            # Clean up notation: fix incomplete measures, pad shorter parts
            for part in score.parts:
                part.makeRests(fillGaps=True, inPlace=True)
                part.makeNotation(inPlace=True)

            maxMeasures = max(
                len(list(p.getElementsByClass('Measure'))) for p in score.parts
            )
            for part in score.parts:
                measures = list(part.getElementsByClass('Measure'))
                if len(measures) < maxMeasures:
                    ts = None
                    for m in reversed(measures):
                        tsList = m.getElementsByClass('TimeSignature')
                        if tsList:
                            ts = tsList[0]
                            break
                    if ts is None:
                        ts = meter.TimeSignature('4/4')
                    for mNum in range(len(measures) + 1, maxMeasures + 1):
                        restMeasure = stream.Measure(number=mNum)
                        r = m21note.Rest(quarterLength=ts.barDuration.quarterLength)
                        restMeasure.append(r)
                        part.append(restMeasure)

            title = self.title_edit.text().strip()
            artist = self.author_edit.text().strip()
            if title:
                score.metadata.title = title
            if artist:
                score.metadata.composer = artist

            score.write('musicxml', fp=path)
            self.log(f"MusicXML saved to: {path}")
            self.log(f"Parts: {len(score.parts)}")
            for i, part in enumerate(score.parts):
                name = part.partName or f"Part {i+1}"
                notes = len(part.flatten().notes)
                self.log(f"  Part {i+1}: {name} ({notes} notes)")
            self.log("")
            self.settings.setValue('last_output_dir', os.path.dirname(path))
        except Exception as e:
            self.log(f"Error: {e}")
            QMessageBox.critical(self, "Export Error", f"Failed to convert MIDI to MusicXML:\n{e}")
        finally:
            self.export_xml_btn.setEnabled(self.file_type == "midi")

    def do_convert(self):
        if not self.midi_path and not self._pr_notes:
            return

        # Suggest a filename: Title - Author.ahk (spaces → underscores)
        title_text = self.title_edit.text().strip()
        author_text = self.author_edit.text().strip()
        if title_text and author_text:
            suggested = f"{title_text.replace(' ', '_')}-{author_text.replace(' ', '_')}.ahk"
        elif title_text:
            suggested = title_text.replace(' ', '_') + '.ahk'
        elif self.midi_path:
            suggested = os.path.splitext(os.path.basename(self.midi_path))[0].replace(' ', '_') + '.ahk'
        else:
            suggested = 'composition.ahk'

        last_dir = self.settings.value('last_output_dir', '')
        if last_dir:
            suggested = os.path.join(last_dir, suggested)

        output, _ = QFileDialog.getSaveFileName(
            self, "Save AHK File", suggested, "AHK Files (*.ahk);;All Files (*)")
        if not output:
            return

        self.settings.setValue('last_output_dir', os.path.dirname(output))

        title = self.title_edit.text().strip() or None
        author = self.author_edit.text().strip() or None

        self.log_text.clear()
        src_name = os.path.basename(self.midi_path) if self.midi_path else "Composition"
        self.log(f"Converting ({self.file_type.upper()}): {src_name}")

        QApplication.processEvents()

        transpose_data = self.transpose_combo.currentData()
        chord_window = self.chord_window_combo.currentData()

        # Use piano roll notes if available (allows editing)
        if self._pr_notes:
            active_notes = self.piano_roll.getActiveNotes()
            if not active_notes:
                QMessageBox.warning(self, "Warning", "No visible notes to convert. Check track visibility.")
                return
            self.log(f"Notes from piano roll: {len(active_notes)} (after edits)")
            inst_data = self.instrument_combo.currentData()
            inst_name = inst_data[0] if inst_data else None
            success, log_lines = convert(self.midi_path or "composition", output, None, title, author,
                                         transpose_data, chord_window_ms=chord_window,
                                         notes_override=active_notes,
                                         use_chords=self.use_chords_cb.isChecked(),
                                         instrument=inst_name,
                                         smooth_octaves=self.smooth_octaves_cb.isChecked())
        elif self.file_type == "musicxml" and self.xml_root is not None:
            xml_notes = extract_notes_musicxml(self.xml_root, None)
            inst_data = self.instrument_combo.currentData()
            inst_name = inst_data[0] if inst_data else None
            success, log_lines = convert(self.midi_path, output, None, title, author,
                                         transpose_data, chord_window_ms=chord_window,
                                         notes_override=xml_notes,
                                         use_chords=self.use_chords_cb.isChecked(),
                                         instrument=inst_name,
                                         smooth_octaves=self.smooth_octaves_cb.isChecked())
        else:
            inst_data = self.instrument_combo.currentData()
            inst_name = inst_data[0] if inst_data else None
            success, log_lines = convert(self.midi_path, output, None, title, author,
                                         transpose_data, chord_window_ms=chord_window,
                                         use_chords=self.use_chords_cb.isChecked(),
                                         instrument=inst_name,
                                         smooth_octaves=self.smooth_octaves_cb.isChecked())

        for line in log_lines:
            self.log(line)

        if success:
            self.log("")
            self.log(f"Done! Output: {output}")
            self._flash_button(self.convert_btn, "Done!", "#4a7")
        else:
            self.log("")
            self.log("ERROR: Conversion failed. Check the log above.")
            self._flash_button(self.convert_btn, "Failed!", "#c44")

    def _update_play_btn_text(self):
        """Update play button text based on current selection state."""
        if self._playback_playing:
            return
        mode = self.piano_roll.getPlaybackMode()
        if mode == 'selection':
            count = sum(1 for n in self._pr_notes if n.selected and not n.deleted)
            self.play_btn.setText(f"▶ Selection ({count})")
        elif mode == 'range':
            lo = min(self.piano_roll._range_start_ms, self.piano_roll._range_end_ms)
            hi = max(self.piano_roll._range_start_ms, self.piano_roll._range_end_ms)
            dur = (hi - lo) / 1000
            self.play_btn.setText(f"▶ Range ({dur:.1f}s)")
        else:
            self.play_btn.setText("▶ Play")

    def _on_gw2_preview_toggled(self, checked):
        """Clear paused playback state so next Play does a fresh render."""
        if self._playback_paused_at > 0:
            self._playback_paused_at = 0
            self._playback_sound = None
            self.playback_slider.setValue(0)
            self._update_time_label()
            self.piano_roll.setCursorMs(-1)
            self._update_play_btn_text()

    def _playback_toggle(self):
        import pygame
        _ensure_pygame()
        if self._playback_playing:
            # Pause
            pygame.mixer.pause()
            self._playback_paused_at = self._playback_elapsed_ms()
            self._playback_playing = False
            self._playback_timer.stop()
            self.play_btn.setText("▶ Play")
            self.gw2_preview_cb.setEnabled(True)
            return

        if self._playback_paused_at > 0 and self._playback_sound:
            # Resume from pause
            pygame.mixer.unpause()
            self._playback_playing = True
            import time
            self._playback_start_time = time.time() * 1000 - self._playback_paused_at
            self._playback_timer.start()
            self.play_btn.setText("⏸ Pause")
            self.gw2_preview_cb.setEnabled(False)
            return

        # Start fresh playback
        if not self._pr_notes:
            return

        self.play_btn.setEnabled(False)
        self.play_btn.setText("Rendering...")

        # Get notes based on selection mode
        active, self._playback_offset_ms = self.piano_roll.getPlaybackNotes()

        if not active:
            self._update_play_btn_text()
            self.play_btn.setEnabled(True)
            return

        self._playback_total_ms = max(a[0] + a[1] for a in active) + 200

        # Render in background thread
        self._render_worker = RenderWorker(active, gw2_preview=self.gw2_preview_cb.isChecked())
        self._render_worker.finished.connect(self._on_render_done)
        self._render_worker.start()

    def _playback_from_here(self):
        """Start playback from the current view position (left edge of scroll)."""
        if self._playback_playing:
            self._playback_stop()

        if not self._pr_notes:
            return

        start_ms = self.piano_roll._scroll_x
        tsc = self.piano_roll._tempo_scale

        # Get all visible notes, then filter to those overlapping or after start_ms
        melody = set(i for i, t in enumerate(self._pr_tracks) if t.get('melody', False))
        visible = set(i for i, t in enumerate(self._pr_tracks) if t.get('visible', True))
        active = []
        for n in self._pr_notes:
            if n.deleted or n.simplified or n.track not in visible:
                continue
            end_ms = n.start_ms + n.duration_ms
            if end_ms <= start_ms:
                continue  # note ends before our start point
            adj_start = max(0, n.start_ms - start_ms) * tsc
            adj_dur = (n.duration_ms if n.start_ms >= start_ms else end_ms - start_ms) * tsc
            active.append((adj_start, adj_dur, n.pitch, n.velocity, n.track in melody))
        active = self.piano_roll._filter_in_range(active)
        active.sort(key=lambda x: x[0])

        if not active:
            return

        self._playback_offset_ms = start_ms * tsc
        self._playback_total_ms = max(a[0] + a[1] for a in active) + 200

        self.play_btn.setEnabled(False)
        self.play_here_btn.setEnabled(False)
        self.play_btn.setText("Rendering...")

        self._render_worker = RenderWorker(active, gw2_preview=self.gw2_preview_cb.isChecked())
        self._render_worker.finished.connect(self._on_render_done)
        self._render_worker.start()

    def _on_render_done(self, audio):
        import pygame
        _ensure_pygame()

        wav = audio_to_wav_bytes(audio)
        self._playback_sound = pygame.mixer.Sound(wav)
        self._playback_sound.play()
        self._playback_playing = True
        self._playback_paused_at = 0

        import time
        self._playback_start_time = time.time() * 1000

        self.play_btn.setText("⏸ Pause")
        self.play_btn.setEnabled(True)
        self.play_here_btn.setEnabled(False)
        self.gw2_preview_cb.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self._playback_timer.start()

    def _playback_stop(self):
        import pygame
        _ensure_pygame()
        pygame.mixer.stop()
        self._playback_playing = False
        self._playback_paused_at = 0
        self._playback_offset_ms = 0
        self._playback_timer.stop()
        self._update_play_btn_text()
        self.play_here_btn.setEnabled(bool(self._pr_notes))
        self.stop_btn.setEnabled(False)
        self.gw2_preview_cb.setEnabled(True)
        self.playback_slider.setValue(0)
        self._update_time_label()
        self.piano_roll.setCursorMs(-1)

    def _playback_elapsed_ms(self):
        if not self._playback_playing:
            return self._playback_paused_at
        import time
        return time.time() * 1000 - self._playback_start_time

    def _playback_tick(self):
        elapsed = self._playback_elapsed_ms()
        if elapsed >= self._playback_total_ms:
            if self._loop_cb.isChecked() and self._playback_sound:
                # Loop: restart audio and timer
                import time
                self._playback_sound.play()
                self._playback_start_time = time.time() * 1000
                elapsed = 0
            else:
                self._playback_stop()
                return
        # Update progress
        progress = int(elapsed / self._playback_total_ms * 1000)
        self.playback_slider.setValue(progress)
        # Update time label (include offset so "Here" shows real song time)
        display_ms = elapsed + self._playback_offset_ms
        total_display_ms = self._playback_total_ms + self._playback_offset_ms
        es = int(display_ms / 1000)
        ts = int(total_display_ms / 1000)
        self.playback_time_label.setText(f"{es // 60}:{es % 60:02d} / {ts // 60}:{ts % 60:02d}")
        # Update piano roll cursor (offset to match original timeline position)
        ts = self.piano_roll._tempo_scale
        self.piano_roll.setCursorMs((elapsed + self._playback_offset_ms) / ts if ts != 1.0 else elapsed + self._playback_offset_ms)
        self.piano_roll.ensureCursorVisible()
        self._minimap.update()

    # ── Drag and drop ─────────────────────────────────────────────
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                path = url.toLocalFile()
                if path.lower().endswith(('.mid', '.midi', '.musicxml', '.mxl', '.ahk')):
                    event.acceptProposedAction()
                    return

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path and os.path.isfile(path):
                self._load_file(path)
                break

    # ── Recent files ──────────────────────────────────────────────
    def _rebuild_recent_menu(self):
        self._recent_menu.clear()
        recents = self.settings.value('recent_files', [])
        if not recents:
            act = self._recent_menu.addAction("(none)")
            act.setEnabled(False)
            return
        for path in recents[:10]:
            act = self._recent_menu.addAction(os.path.basename(path))
            act.setToolTip(path)
            act.triggered.connect(lambda checked, p=path: self._load_file(p))
        self._recent_menu.addSeparator()
        self._recent_menu.addAction("Clear Recent", lambda: (self.settings.setValue('recent_files', []), self._rebuild_recent_menu()))

    def _add_recent_file(self, path):
        recents = self.settings.value('recent_files', [])
        if not isinstance(recents, list):
            recents = []
        if path in recents:
            recents.remove(path)
        recents.insert(0, path)
        self.settings.setValue('recent_files', recents[:10])
        self._rebuild_recent_menu()

    # ── Stats bar ─────────────────────────────────────────────────
    def _update_stats(self):
        active = [n for n in self._pr_notes if not n.deleted]
        if not active:
            self._stats_label.setText("")
            return
        total = len(active)
        dur_ms = max(n.start_ms + n.duration_ms for n in active)
        dur_s = dur_ms / 1000
        pitches = [n.pitch for n in active]
        min_p, max_p = min(pitches), max(pitches)
        octave_span = (max_p - min_p) // 12
        # Count octave changes needed (simplified)
        sorted_notes = sorted(active, key=lambda n: n.start_ms)
        oct_changes = 0
        inst_data = self.instrument_combo.currentData()
        if inst_data:
            oct_size = 12
            prev_oct = -1
            for n in sorted_notes:
                cur_oct = (n.pitch - inst_data[2]) // oct_size
                if prev_oct >= 0 and cur_oct != prev_oct:
                    oct_changes += 1
                prev_oct = cur_oct
        out_of_range = 0
        if inst_data:
            out_of_range = sum(1 for n in active if n.pitch < inst_data[2] or n.pitch > inst_data[3])
        self._stats_label.setText(
            f"Notes: {total}  |  Duration: {dur_s:.1f}s  |  "
            f"Octave span: {octave_span}  |  Octave changes: ~{oct_changes}  |  "
            f"Out of range: {out_of_range}")
        if hasattr(self, '_minimap'):
            self._minimap.update()

    # ── Click-to-seek (progress bar + minimap) ────────────────────
    def _on_progress_click(self, event):
        if self._playback_total_ms <= 0:
            return
        ratio = event.position().x() / self.playback_slider.width()
        ratio = max(0, min(1, ratio))
        ms = ratio * self._playback_total_ms
        self._seek_to_ms(ms + self._playback_offset_ms)

    def _on_minimap_seek(self, ms):
        # Scroll the piano roll to this position
        self.piano_roll._scroll_x = max(0, ms)
        self.piano_roll._update_scrollbars()
        self.piano_roll.update()
        self._minimap.update()

    def _update_time_label(self):
        """Update the time label to show current viewport position / total duration."""
        if self._playback_playing:
            return  # During playback, the playback timer handles the label
        pr = self.piano_roll
        scale = pr._tempo_scale
        current_ms = pr._scroll_x * scale
        total_ms = pr._total_ms * scale
        cs = int(current_ms / 1000)
        ts = int(total_ms / 1000)
        self.playback_time_label.setText(f"{cs // 60}:{cs % 60:02d} / {ts // 60}:{ts % 60:02d}")

    def _seek_to_ms(self, ms):
        """Seek playback to a specific ms position."""
        self.piano_roll.setCursorMs(ms)
        self.piano_roll.ensureCursorVisible()

    # ── Toggle methods ────────────────────────────────────────────
    def _toggle_snap(self, checked):
        if checked:
            bpm = 120
            if hasattr(self.piano_roll, '_bpm') and self.piano_roll._bpm:
                bpm = self.piano_roll._bpm
            self.piano_roll._snap_ms = 60000.0 / bpm / 4  # snap to 1/16th notes
        else:
            self.piano_roll._snap_ms = 0.0

    def _on_speed_changed(self, value):
        """Update tempo scale when speed spinbox changes."""
        self.piano_roll._tempo_scale = 100.0 / value
        self._update_time_label()
        self.log(f"Speed: {value}%")

    def _set_theme(self, theme):
        app = QApplication.instance()
        if theme == 'light':
            from PyQt6.QtGui import QPalette
            pal = QPalette()
            app.setPalette(pal)
            app.setStyleSheet("")
        else:
            app.setStyle("Fusion")
            from PyQt6.QtGui import QPalette
            pal = QPalette()
            pal.setColor(QPalette.ColorRole.Window, QColor(53, 53, 53))
            pal.setColor(QPalette.ColorRole.WindowText, QColor(255, 255, 255))
            pal.setColor(QPalette.ColorRole.Base, QColor(35, 35, 35))
            pal.setColor(QPalette.ColorRole.AlternateBase, QColor(53, 53, 53))
            pal.setColor(QPalette.ColorRole.ToolTipBase, QColor(25, 25, 25))
            pal.setColor(QPalette.ColorRole.ToolTipText, QColor(255, 255, 255))
            pal.setColor(QPalette.ColorRole.Text, QColor(255, 255, 255))
            pal.setColor(QPalette.ColorRole.Button, QColor(53, 53, 53))
            pal.setColor(QPalette.ColorRole.ButtonText, QColor(255, 255, 255))
            pal.setColor(QPalette.ColorRole.BrightText, QColor(255, 0, 0))
            pal.setColor(QPalette.ColorRole.Link, QColor(42, 130, 218))
            pal.setColor(QPalette.ColorRole.Highlight, QColor(42, 130, 218))
            pal.setColor(QPalette.ColorRole.HighlightedText, QColor(35, 35, 35))
            app.setPalette(pal)

    # ── Tools: track operations ───────────────────────────────────
    def _merge_tracks(self):
        """Merge all checked tracks into the first checked track."""
        checked = []
        for i in range(self.track_list.count()):
            item = self.track_list.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                checked.append(item.data(Qt.ItemDataRole.UserRole))
        if len(checked) < 2:
            self.log("Check at least 2 tracks to merge.")
            return
        self.piano_roll.pushUndo()
        target = checked[0]
        merged = 0
        for n in self._pr_notes:
            if n.track in checked[1:] and not n.deleted:
                n.track = target
                merged += 1
        self.piano_roll.notesChanged.emit()
        self.piano_roll.update()
        self.log(f"Merged {merged} notes from {len(checked)-1} tracks into '{self._pr_tracks[target]['name']}'")
        self._on_notes_changed()

    def _split_track_by_pitch(self):
        """Split selected track into upper and lower halves by median pitch."""
        item = self.track_list.currentItem()
        if item is None:
            self.log("Select a track first.")
            return
        idx = item.data(Qt.ItemDataRole.UserRole)
        track_notes = [n for n in self._pr_notes if n.track == idx and not n.deleted]
        if not track_notes:
            self.log("No notes in selected track.")
            return
        self.piano_roll.pushUndo()
        median = sorted(n.pitch for n in track_notes)[len(track_notes) // 2]
        new_idx = len(self._pr_tracks)
        color = TRACK_COLORS[new_idx % len(TRACK_COLORS)]
        self._pr_tracks.append({
            'index': new_idx,
            'name': f"{self._pr_tracks[idx]['name']} (high)",
            'visible': True,
            'melody': False,
            'simplify': False,
            'color': color,
            'notes': 0,
        })
        moved = 0
        for n in track_notes:
            if n.pitch >= median:
                n.track = new_idx
                moved += 1
        self.piano_roll.notesChanged.emit()
        self.piano_roll.update()
        self.log(f"Split track: {moved} notes (≥{median}) moved to new track")
        self._on_notes_changed()

    def _smart_octave(self):
        """Assign each track to the octave that minimizes out-of-range notes."""
        if not self._pr_notes:
            self.log("No notes loaded.")
            return
        inst_data = self.instrument_combo.currentData()
        if not inst_data:
            return
        lo, hi = inst_data[2], inst_data[3]
        self.piano_roll.pushUndo()
        total_shifted = 0
        for tidx, t in enumerate(self._pr_tracks):
            track_notes = [n for n in self._pr_notes if n.track == tidx and not n.deleted]
            if not track_notes:
                continue
            best_shift = 0
            best_oor = len(track_notes)
            for shift in range(-48, 49, 12):
                oor = sum(1 for n in track_notes if (n.pitch + shift) < lo or (n.pitch + shift) > hi)
                if oor < best_oor:
                    best_oor = oor
                    best_shift = shift
            if best_shift != 0:
                for n in track_notes:
                    n.pitch += best_shift
                total_shifted += len(track_notes)
                direction = "up" if best_shift > 0 else "down"
                self.log(f"  {t['name']}: shifted {abs(best_shift)//12} octaves {direction} ({best_oor} remaining out-of-range)")
        self.piano_roll.notesChanged.emit()
        self.piano_roll.update()
        self.log(f"Smart octave: adjusted {total_shifted} notes across {len(self._pr_tracks)} tracks")

    # ── Song Analysis ─────────────────────────────────────────────
    def _analyse_song(self):
        """Run comprehensive GW2 playback analysis and show report dialog."""
        has_melody = any(t.get('melody', False) for t in self._pr_tracks)
        if not has_melody:
            # Prompt user to pick melody + preserved tracks before analysing
            if not self._pr_tracks:
                return
            from PyQt6.QtWidgets import QRadioButton, QButtonGroup
            dlg = QDialog(self)
            dlg.setWindowTitle("Select Tracks")
            layout = QVBoxLayout(dlg)
            layout.addWidget(QLabel("Which track is the melody?"))
            btn_group = QButtonGroup(dlg)
            preserve_checks = []
            for i, t in enumerate(self._pr_tracks):
                r, g, b = TRACK_COLORS[i % len(TRACK_COLORS)]
                row = QHBoxLayout()
                rb = QRadioButton(t['name'])
                rb.setStyleSheet(f"QRadioButton {{ color: rgb({r},{g},{b}); }}")
                if i == 0:
                    rb.setChecked(True)
                btn_group.addButton(rb, i)
                row.addWidget(rb)
                cb = QCheckBox("Preserve")
                cb.setStyleSheet(f"QCheckBox {{ color: rgb({r},{g},{b}); }}")
                cb.setChecked(t.get('preserve', False))
                preserve_checks.append(cb)
                row.addWidget(cb)
                layout.addLayout(row)
            btns = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
            btns.accepted.connect(dlg.accept)
            btns.rejected.connect(dlg.reject)
            layout.addWidget(btns)
            if dlg.exec() != QDialog.DialogCode.Accepted:
                return
            mel_idx = btn_group.checkedId()
            self._set_melody_track(mel_idx)
            for i, cb in enumerate(preserve_checks):
                if i != mel_idx:
                    self._pr_tracks[i]['preserve'] = cb.isChecked()
                else:
                    self._pr_tracks[i]['preserve'] = False
            self._populate_track_list()
        inst_data = self.instrument_combo.currentData()
        if not inst_data:
            return
        inst_lo, inst_hi = inst_data[2], inst_data[3]

        report = analyze_song(self._pr_notes, self._pr_tracks, inst_lo, inst_hi)
        if not report:
            QMessageBox.information(self, "Analysis", "No notes to analyse.")
            return

        dlg = AnalysisDialog(report, self)
        dlg.exec()
        if dlg.auto_fix_accepted:
            self._auto_fix_song(report)

    def _auto_fix_song(self, report):
        """Apply suggested fixes from an analysis report.

        Phase 1 – Relative-positioning octave shifts:
          • Melody is shifted first (up only, never down unless entirely above range).
          • Non-melody tracks are then shifted to preserve their original pitch
            offset from melody while minimizing out-of-range notes.  If a non-melody
            track would land in the same GW2 octave as melody, it prefers going down.
        Phase 2 – Remaining fixes:
          • Non-melody: clamp, debass, simplify.
          • All tracks: timing trim.
        """
        self.piano_roll.pushUndo()
        inst_data = self.instrument_combo.currentData()
        if not inst_data:
            return
        inst_lo, inst_hi = inst_data[2], inst_data[3]
        base_pitch = inst_lo + 12

        fix_log = []

        def _median_pitch(notes_list):
            pitches = sorted(n.pitch for n in notes_list)
            return pitches[len(pitches) // 2] if pitches else None

        def _best_shift_oor(tnotes, shift_range):
            """Find the shift in shift_range that minimizes OOR notes."""
            best_s, best_o = 0, len(tnotes)
            for s in shift_range:
                o = sum(1 for n in tnotes
                        if (n.pitch + s) < inst_lo or (n.pitch + s) > inst_hi)
                if o < best_o:
                    best_o = o
                    best_s = s
            return best_s, best_o

        # ── Phase 1: Octave shifts with relative positioning ──────

        # Snapshot original median pitches before any shifts
        orig_medians = {}  # tidx → median pitch
        for tr in report['tracks']:
            tnotes = [n for n in self._pr_notes
                      if n.track == tr['index'] and not n.deleted and not n.simplified]
            m = _median_pitch(tnotes)
            if m is not None:
                orig_medians[tr['index']] = m

        # 1a. Shift melody track first.
        # Factor in cross-octave cost: shifting melody to a different GW2
        # octave from accompaniment doubles octave switches in the AHK output.
        # Score each candidate shift as: OOR_notes + cross_octave_notes * 0.3
        melody_shift = 0
        melody_tidx = None
        for tr in report['tracks']:
            if not tr['is_melody']:
                continue
            melody_tidx = tr['index']
            if 'smart_octave' not in tr['fixes']:
                break
            tnotes = [n for n in self._pr_notes
                      if n.track == melody_tidx and not n.deleted and not n.simplified]
            if not tnotes:
                break

            # Collect non-melody notes for cross-octave estimation
            non_mel_notes = [n for n in self._pr_notes
                             if n.track != melody_tidx and not n.deleted
                             and not n.simplified]

            best_shift = 0
            best_score = float('inf')
            for shift in range(-48, 49, 12):
                # Only allow downward shift if ALL notes are above range
                if shift < 0 and not all(n.pitch > inst_hi for n in tnotes):
                    continue
                oor = sum(1 for n in tnotes
                          if (n.pitch + shift) < inst_lo or (n.pitch + shift) > inst_hi)
                # Estimate cross-octave cost: how many non-melody notes
                # would be in a different GW2 octave from the shifted melody
                mel_median = _median_pitch(tnotes)
                if mel_median is not None and non_mel_notes:
                    mel_oct = _note_octave(mel_median + shift, base_pitch)
                    cross = sum(1 for n in non_mel_notes
                                if _note_octave(n.pitch, base_pitch) != mel_oct)
                else:
                    cross = 0
                score = oor + cross * 0.3
                # Among equal scores, prefer higher position
                if score < best_score or (score == best_score and shift > best_shift):
                    best_score = score
                    best_shift = shift
            if best_shift != 0:
                for n in tnotes:
                    n.pitch += best_shift
                melody_shift = best_shift
                direction = "up" if best_shift > 0 else "down"
                fix_log.append(f"  {tr['name']}: shifted {abs(best_shift)//12} octave(s) {direction}")
            break

        melody_median_new = None
        if melody_tidx is not None and melody_tidx in orig_medians:
            melody_median_new = orig_medians[melody_tidx] + melody_shift

        # 1b. Shift non-melody tracks preserving relative position to melody.
        # The original arrangement is more important than minimising OOR —
        # accompaniment should stay below melody even if some notes go OOR.
        for tr in report['tracks']:
            if tr['is_melody'] or 'smart_octave' not in tr['fixes']:
                continue
            tidx = tr['index']
            name = tr['name']
            tnotes = [n for n in self._pr_notes
                      if n.track == tidx and not n.deleted and not n.simplified]
            if not tnotes:
                continue

            track_median = orig_medians.get(tidx)
            if track_median is None:
                continue

            # Compute ideal shift: preserve original offset from melody
            if melody_median_new is not None and melody_tidx in orig_medians:
                orig_offset = track_median - orig_medians[melody_tidx]
                target_median = melody_median_new + orig_offset
                ideal_shift = round((target_median - track_median) / 12) * 12
            else:
                ideal_shift = 0

            # Use the ideal shift if it keeps at least half the notes in range.
            # Only fall back to nearby alternatives if ideal is really bad.
            def _oor(shift):
                return sum(1 for n in tnotes
                           if (n.pitch + shift) < inst_lo or (n.pitch + shift) > inst_hi)

            ideal_oor = _oor(ideal_shift)
            best_shift = ideal_shift

            if ideal_oor > len(tnotes) * 0.5:
                # Ideal is too lossy — try nearby shifts, prefer closest to ideal
                for delta in [12, -12, 24, -24, 36, -36, 48, -48]:
                    alt = ideal_shift + delta
                    if -48 <= alt <= 48:
                        alt_oor = _oor(alt)
                        if alt_oor < ideal_oor:
                            best_shift = alt
                            ideal_oor = alt_oor
                            if ideal_oor <= len(tnotes) * 0.5:
                                break

            if best_shift != 0:
                for n in tnotes:
                    n.pitch += best_shift
                direction = "up" if best_shift > 0 else "down"
                oor_count = _oor(0)  # OOR after shift (shift already applied)
                oor_note = f" ({sum(1 for n in tnotes if n.pitch < inst_lo or n.pitch > inst_hi)} OOR)" if any(n.pitch < inst_lo or n.pitch > inst_hi for n in tnotes) else ""
                fix_log.append(f"  {name}: shifted {abs(best_shift)//12} octave(s) {direction}{oor_note}")

        # ── Phase 1.5: Harmonic substitution for OOR melody notes ──
        # For melody notes just above range (1-12 semitones over inst_hi),
        # drop one octave and add a perfect fifth to imply the brightness
        # of the original pitch.  Both notes must fit in range and same
        # GW2 octave.
        if melody_tidx is not None:
            mel_oor = [n for n in self._pr_notes
                       if n.track == melody_tidx and not n.deleted
                       and not n.simplified
                       and n.pitch > inst_hi
                       and n.pitch <= inst_hi + 12]
            CHORD_WIN_HS = 5  # ms tolerance for simultaneous
            harm_count = 0
            new_notes = []
            for n in mel_oor:
                dropped = n.pitch - 12
                fifth = dropped + 7
                if dropped < inst_lo or fifth > inst_hi:
                    continue
                # Both must be in the same GW2 octave
                if _note_octave(dropped, base_pitch) != _note_octave(fifth, base_pitch):
                    continue
                # Check chord density — don't exceed 4 simultaneous notes
                sim = sum(1 for x in self._pr_notes
                          if abs(x.start_ms - n.start_ms) <= CHORD_WIN_HS
                          and not x.deleted and not x.simplified and x is not n)
                if sim >= 3:  # adding 2 notes (dropped + fifth) to 3 existing = 5, too many
                    continue
                # Replace original with dropped pitch, add fifth as new note
                n.pitch = dropped
                harm_note = PianoRollNote(n.start_ms, n.duration_ms, fifth,
                                          n.velocity, n.track)
                new_notes.append(harm_note)
                harm_count += 1
            if new_notes:
                self._pr_notes.extend(new_notes)
            if harm_count:
                fix_log.append(f"  Melody: {harm_count} notes substituted (octave down + fifth)")

        # ── Phase 1.75: Octave consolidation for rapid melody runs ──
        # When melody rapidly alternates between two GW2 octaves, snap
        # minority notes into the dominant octave of the run.  This
        # eliminates unnecessary octave switching in fast passages.
        CONSOLIDATE_WIN_MS = 1500  # sliding window for run detection
        CONSOLIDATE_MIN_NOTES = 4  # minimum notes to consider a "run"
        CONSOLIDATE_MIN_CROSSINGS = 2  # min octave crossings to trigger
        if melody_tidx is not None:
            mel_notes_sorted = sorted(
                [n for n in self._pr_notes
                 if n.track == melody_tidx and not n.deleted and not n.simplified],
                key=lambda n: n.start_ms)
            consolidated = 0
            i = 0
            while i < len(mel_notes_sorted):
                # Find the end of a window starting at note i
                j = i + 1
                while j < len(mel_notes_sorted) and mel_notes_sorted[j].start_ms - mel_notes_sorted[i].start_ms <= CONSOLIDATE_WIN_MS:
                    j += 1
                run = mel_notes_sorted[i:j]
                if len(run) >= CONSOLIDATE_MIN_NOTES:
                    # Count octave crossings in this run
                    octs = [_note_octave(n.pitch, base_pitch) for n in run]
                    crossings = sum(1 for k in range(1, len(octs)) if octs[k] != octs[k - 1])
                    if crossings >= CONSOLIDATE_MIN_CROSSINGS:
                        # Find dominant octave
                        oct_counts = Counter(octs)
                        dominant_oct = oct_counts.most_common(1)[0][0]
                        # Snap minority notes into dominant octave
                        for n in run:
                            n_oct = _note_octave(n.pitch, base_pitch)
                            if n_oct != dominant_oct:
                                if n_oct < dominant_oct:
                                    new_pitch = n.pitch + 12
                                else:
                                    new_pitch = n.pitch - 12
                                # Only snap if the new pitch is in instrument range
                                if inst_lo <= new_pitch <= inst_hi:
                                    n.pitch = new_pitch
                                    consolidated += 1
                    i = j  # skip past this run
                else:
                    i += 1
            if consolidated:
                fix_log.append(f"  Melody: {consolidated} notes consolidated into dominant octave")

        # ── Phase 1.9: Cross-octave accompaniment cleanup ─────────
        # Accompaniment notes in a different GW2 octave than simultaneous
        # melody notes force an octave switch that can confuse the player.
        # Try to shift them into the melody's octave; delete if impossible.
        CHORD_WIN_XO = 5  # ms tolerance for "simultaneous"
        if melody_tidx is not None:
            mel_sorted_xo = sorted(
                [n for n in self._pr_notes
                 if n.track == melody_tidx and not n.deleted and not n.simplified],
                key=lambda n: n.start_ms)
            mel_starts_xo = [n.start_ms for n in mel_sorted_xo]
            # Build a lookup: for a given time, what GW2 octave is the melody in?
            # If multiple melody notes at same time, use the most common octave.
            shifted_xo = 0
            simplified_xo = 0
            non_mel = sorted(
                [n for n in self._pr_notes
                 if n.track != melody_tidx and not n.deleted and not n.simplified],
                key=lambda n: n.start_ms)
            mi_xo = 0
            for n in non_mel:
                # Find melody notes simultaneous with this note
                while mi_xo < len(mel_starts_xo) - 1 and mel_starts_xo[mi_xo] < n.start_ms - CHORD_WIN_XO:
                    mi_xo += 1
                mel_octs = []
                for k in range(max(0, mi_xo - 1), min(len(mel_sorted_xo), mi_xo + 3)):
                    if abs(mel_sorted_xo[k].start_ms - n.start_ms) <= CHORD_WIN_XO:
                        mel_octs.append(_note_octave(mel_sorted_xo[k].pitch, base_pitch))
                if not mel_octs:
                    continue
                mel_oct = max(set(mel_octs), key=mel_octs.count)
                n_oct = _note_octave(n.pitch, base_pitch)
                if n_oct == mel_oct:
                    continue
                # Try shifting ±12 to match melody octave
                fixed = False
                for delta in (12, -12, 24, -24):
                    new_pitch = n.pitch + delta
                    if inst_lo <= new_pitch <= inst_hi and _note_octave(new_pitch, base_pitch) == mel_oct:
                        n.pitch = new_pitch
                        shifted_xo += 1
                        fixed = True
                        break
                if not fixed:
                    is_preserved = self._pr_tracks[n.track].get('preserve', False)
                    if not is_preserved:
                        n.simplified = True
                        simplified_xo += 1
            if shifted_xo or simplified_xo:
                parts = []
                if shifted_xo:
                    parts.append(f"{shifted_xo} shifted")
                if simplified_xo:
                    parts.append(f"{simplified_xo} simplified")
                fix_log.append(f"  Cross-octave cleanup: {', '.join(parts)}")

        # ── Phase 2: Remaining fixes ──────────────────────────────

        for tr in report['tracks']:
            tidx = tr['index']
            is_mel = tr['is_melody']
            fixes = tr['fixes']
            name = tr['name']

            # Clamp remaining out-of-range — non-melody (incl. preserved)
            if 'clamp' in fixes and not is_mel:
                tnotes = [n for n in self._pr_notes
                          if n.track == tidx and not n.deleted and not n.simplified]
                count = 0
                for n in tnotes:
                    if n.pitch < inst_lo:
                        n.pitch = inst_lo + (n.pitch % 12)
                        if n.pitch < inst_lo:
                            n.pitch += 12
                        count += 1
                    elif n.pitch > inst_hi:
                        target = inst_lo + (n.pitch % 12)
                        while target + 12 <= inst_hi:
                            target += 12
                        n.pitch = target
                        count += 1
                if count:
                    fix_log.append(f"  {name}: clamped {count} out-of-range notes")

            # Bass duplicate removal — non-melody, non-preserved only
            is_preserved = self._pr_tracks[tidx].get('preserve', False)
            if 'debass' in fixes and not is_mel and not is_preserved:
                if not self._pr_tracks[tidx].get('debass', False):
                    self._pr_tracks[tidx]['debass'] = True
                    fix_log.append(f"  {name}: enabled bass removal")

            # Simplify — non-melody, non-preserved only
            if 'simplify' in fixes and not is_mel and not is_preserved:
                if not self._pr_tracks[tidx].get('simplify', False):
                    self._pr_tracks[tidx]['simplify'] = True
                    fix_log.append(f"  {name}: enabled simplify")

        # ── Phase 2.5: Density-adaptive thinning ─────────────────
        # Dynamically adjust the max chord size at each beat based on
        # local melody density.  Sparse melody → rich chords allowed.
        # Dense melody → accompaniment thinned or muted.
        # Priority: melody > preserved > other accompaniment.
        DENSITY_HALF_WIN = 500  # ms half-window for melody density
        CHORD_WIN_THIN = 5     # ms tolerance for "simultaneous"
        if melody_tidx is not None:
            protected_tracks = set()
            protected_tracks.add(melody_tidx)
            for t in self._pr_tracks:
                if t.get('preserve', False):
                    protected_tracks.add(t['index'])

            # Build sorted melody start times + pitches for density lookup
            mel_sorted = sorted(
                [(n.start_ms, n.pitch) for n in self._pr_notes
                 if n.track == melody_tidx and not n.deleted and not n.simplified],
                key=lambda x: x[0])
            mel_starts = [x[0] for x in mel_sorted]
            mel_pitches = [x[1] for x in mel_sorted]

            def _local_max_chord(time_ms):
                """Max chord size based on melody density + octave crossings."""
                lo = bisect_left(mel_starts, time_ms - DENSITY_HALF_WIN)
                hi = bisect_right(mel_starts, time_ms + DENSITY_HALF_WIN)
                note_count = hi - lo
                # Count octave transitions in window — each costs ~60ms overhead
                oct_changes = 0
                for k in range(lo + 1, hi):
                    if _note_octave(mel_pitches[k], base_pitch) != _note_octave(mel_pitches[k - 1], base_pitch):
                        oct_changes += 1
                effective_density = note_count + oct_changes
                if effective_density <= 2:
                    return 4  # sparse: full chords
                elif effective_density <= 5:
                    return 3  # moderate: slightly thinner
                else:
                    return 1  # dense/cross-octave: melody only

            # Group active notes by beat time
            active = [n for n in self._pr_notes
                      if not n.deleted and not n.simplified]
            active.sort(key=lambda n: n.start_ms)
            muted_count = 0
            i = 0
            while i < len(active):
                # Collect simultaneous notes
                group = [active[i]]
                j = i + 1
                while j < len(active) and active[j].start_ms - active[i].start_ms <= CHORD_WIN_THIN:
                    group.append(active[j])
                    j += 1
                max_chord = _local_max_chord(active[i].start_ms)
                if len(group) > max_chord:
                    # Sort: melody first, then preserved, then others
                    def _priority(n):
                        if n.track == melody_tidx:
                            return 0
                        if n.track in protected_tracks:
                            return 1
                        return 2
                    group.sort(key=lambda n: (_priority(n), -n.pitch))
                    # Delete excess notes (lowest priority first, from end)
                    for n in group[max_chord:]:
                        if n.track not in protected_tracks:
                            n.deleted = True
                            muted_count += 1
                i = j
            if muted_count:
                fix_log.append(f"  Thinned {muted_count} accompaniment notes (density-adaptive)")

        for tr in report['tracks']:
            tidx = tr['index']
            is_mel = tr['is_melody']
            fixes = tr['fixes']
            name = tr['name']

            # Timing fixes — safe for all tracks including melody
            if 'timing' in fixes:
                tnotes = sorted(
                    [n for n in self._pr_notes
                     if n.track == tidx and not n.deleted and not n.simplified],
                    key=lambda n: n.start_ms)
                count = 0
                for i in range(len(tnotes) - 1):
                    gap = tnotes[i + 1].start_ms - (tnotes[i].start_ms + tnotes[i].duration_ms)
                    co = _note_octave(tnotes[i].pitch, base_pitch)
                    no = _note_octave(tnotes[i + 1].pitch, base_pitch)
                    min_gap = GW2_MIN_NOTE_DELAY_MS + (GW2_OCTAVE_SWAP_DELAY_MS if co != no else 0)
                    if 0 <= gap < min_gap:
                        new_dur = tnotes[i + 1].start_ms - tnotes[i].start_ms - min_gap
                        if new_dur < 10:
                            new_dur = 10
                        if new_dur != tnotes[i].duration_ms:
                            tnotes[i].duration_ms = new_dur
                            count += 1
                if count:
                    fix_log.append(f"  {name}: adjusted timing on {count} notes")

        # Update UI
        self._populate_track_list()
        self.piano_roll.updateSimplifiedNotes()
        self.piano_roll.notesChanged.emit()
        self.piano_roll.update()

        if fix_log:
            self.log("Auto-fix applied:\n" + "\n".join(fix_log))
        else:
            self.log("Auto-fix: no changes needed")

    # ── Tools: note operations ────────────────────────────────────
    def _remove_duplicates(self):
        """Remove overlapping notes with identical pitch and start time."""
        if not self._pr_notes:
            return
        self.piano_roll.pushUndo()
        seen = set()
        removed = 0
        for n in sorted(self._pr_notes, key=lambda x: (x.start_ms, x.pitch)):
            if n.deleted:
                continue
            key = (round(n.start_ms, 1), n.pitch)
            if key in seen:
                n.deleted = True
                removed += 1
            else:
                seen.add(key)
        self.piano_roll.notesChanged.emit()
        self.piano_roll.update()
        self.log(f"Removed {removed} duplicate notes")

    def _merge_short_gaps(self):
        """Merge notes separated by tiny gaps (< 50ms) into longer notes."""
        if not self._pr_notes:
            return
        self.piano_roll.pushUndo()
        active = sorted([n for n in self._pr_notes if not n.deleted], key=lambda x: (x.pitch, x.track, x.start_ms))
        merged = 0
        i = 0
        while i < len(active) - 1:
            a, b = active[i], active[i + 1]
            if a.pitch == b.pitch and a.track == b.track:
                gap = b.start_ms - (a.start_ms + a.duration_ms)
                if 0 < gap < 50:
                    a.duration_ms = (b.start_ms + b.duration_ms) - a.start_ms
                    b.deleted = True
                    merged += 1
                    continue
            i += 1
        self.piano_roll.notesChanged.emit()
        self.piano_roll.update()
        self.log(f"Merged {merged} short gaps")

    def _edit_velocity(self):
        """Edit velocity of selected notes via dialog."""
        sel = [n for n in self._pr_notes if n.selected and not n.deleted]
        if not sel:
            self.log("Select notes first.")
            return
        from PyQt6.QtWidgets import QInputDialog
        val, ok = QInputDialog.getInt(self, "Edit Velocity", f"Velocity (1-127) for {len(sel)} notes:", sel[0].velocity, 1, 127)
        if ok:
            self.piano_roll.pushUndo()
            for n in sel:
                n.velocity = val
            self.piano_roll.update()
            self.log(f"Set velocity to {val} for {len(sel)} notes")

    # ── Export MIDI ───────────────────────────────────────────────
    def _export_midi(self):
        """Export piano roll notes as a MIDI file."""
        if not self._pr_notes:
            QMessageBox.warning(self, "Warning", "No notes to export.")
            return
        import mido
        path, _ = QFileDialog.getSaveFileName(self, "Export MIDI", "", "MIDI Files (*.mid);;All Files (*)")
        if not path:
            return
        mid = mido.MidiFile(ticks_per_beat=480)
        bpm = 120
        us_per_beat = int(60_000_000 / bpm)
        ticks_per_ms = 480 / (us_per_beat / 1000)
        # Group notes by track
        track_groups = {}
        for n in self._pr_notes:
            if not n.deleted:
                track_groups.setdefault(n.track, []).append(n)
        for tidx in sorted(track_groups):
            trk = mido.MidiTrack()
            mid.tracks.append(trk)
            if tidx == 0:
                trk.append(mido.MetaMessage('set_tempo', tempo=us_per_beat))
            if tidx < len(self._pr_tracks):
                trk.append(mido.MetaMessage('track_name', name=self._pr_tracks[tidx]['name']))
            events = []
            for n in track_groups[tidx]:
                t_on = int(n.start_ms * ticks_per_ms)
                t_off = int((n.start_ms + n.duration_ms) * ticks_per_ms)
                events.append((t_on, 'note_on', n.pitch, n.velocity))
                events.append((t_off, 'note_off', n.pitch, 0))
            events.sort(key=lambda e: e[0])
            prev_tick = 0
            for tick, msg_type, pitch, vel in events:
                delta = tick - prev_tick
                trk.append(mido.Message(msg_type, note=pitch, velocity=vel, time=delta))
                prev_tick = tick
        mid.save(path)
        self.log(f"Exported MIDI to {os.path.basename(path)}")

    # ── Batch Convert ─────────────────────────────────────────────
    def _batch_convert(self):
        """Select multiple MIDI files and convert them all to AHK."""
        files, _ = QFileDialog.getOpenFileNames(self, "Select MIDI Files", "",
            "MIDI Files (*.mid *.midi);;All Files (*)")
        if not files:
            return
        out_dir = QFileDialog.getExistingDirectory(self, "Select Output Directory")
        if not out_dir:
            return
        transpose_data = self.transpose_combo.currentData()
        chord_window = self.chord_window_combo.currentData()
        inst_data = self.instrument_combo.currentData()
        inst_name = inst_data[0] if inst_data else None
        success_count = 0
        for fpath in files:
            basename = os.path.splitext(os.path.basename(fpath))[0]
            outpath = os.path.join(out_dir, basename + '.ahk')
            try:
                ok, _ = convert(fpath, outpath, None, basename, None,
                               transpose_data, chord_window_ms=chord_window,
                               use_chords=self.use_chords_cb.isChecked(),
                               instrument=inst_name,
                               smooth_octaves=self.smooth_octaves_cb.isChecked())
                if ok:
                    success_count += 1
                    self.log(f"  ✓ {basename}")
                else:
                    self.log(f"  ✗ {basename} (failed)")
            except Exception as e:
                self.log(f"  ✗ {basename}: {e}")
        self.log(f"Batch convert: {success_count}/{len(files)} files converted")

    # ── Session save / restore ─────────────────────────────────────

    def _session_path(self):
        return os.path.join(os.path.expanduser('~'), '.serenade_session.json')

    def _save_session(self):
        """Save current editing state for next session."""
        import json
        def _note_to_dict(n):
            d = {'s': n.start_ms, 'd': n.duration_ms, 'p': n.pitch,
                 'v': n.velocity, 't': n.track, 'sel': n.selected,
                 'del': n.deleted, 'w': n.warning}
            if n.simplified_manual is not None:
                d['sm'] = n.simplified_manual
            return d
        def _notes_list(notes):
            return [_note_to_dict(n) for n in notes]

        session = {
            'midi_path': self.midi_path,
            'file_type': self.file_type,
            'title': self.title_edit.text(),
            'artist': self.author_edit.text(),
            'instrument_idx': self.instrument_combo.currentIndex(),
            'transpose_idx': self.transpose_combo.currentIndex(),
            'chord_window_idx': self.chord_window_combo.currentIndex(),
            'use_chords': self.use_chords_cb.isChecked(),
            'notes': _notes_list(self._pr_notes),
            'tracks': self._pr_tracks,
            'scroll_x': self.piano_roll._scroll_x,
            'scroll_y': self.piano_roll._scroll_y,
            'px_per_ms': self.piano_roll._px_per_ms,
            'total_ms': self.piano_roll._total_ms,
            'gw2_only': self.piano_roll._gw2_only,
            'pitch_min': self.piano_roll._pitch_min,
            'pitch_max': self.piano_roll._pitch_max,
            'gw2_pitch_min': self.piano_roll.GW2_PITCH_MIN,
            'gw2_pitch_max': self.piano_roll.GW2_PITCH_MAX,
        }
        try:
            with open(self._session_path(), 'w') as f:
                json.dump(session, f, separators=(',', ':'))
        except Exception:
            pass

    def _restore_session(self):
        """Restore editing state from previous session."""
        import json
        path = self._session_path()
        if not os.path.exists(path):
            return
        try:
            with open(path, 'r') as f:
                session = json.load(f)
        except Exception:
            return

        def _dict_to_note(d):
            n = PianoRollNote(d['s'], d['d'], d['p'], d['v'], d['t'])
            n.selected = d.get('sel', False)
            n.deleted = d.get('del', False)
            n.warning = d.get('w', '')
            n.simplified_manual = d.get('sm', None)
            return n

        notes = [_dict_to_note(d) for d in session.get('notes', [])]
        if not notes:
            return

        self.midi_path = session.get('midi_path', '')
        self.file_type = session.get('file_type', 'midi')
        self.title_edit.setText(session.get('title', ''))
        self.author_edit.setText(session.get('artist', ''))
        self.instrument_combo.setCurrentIndex(session.get('instrument_idx', 0))
        self.transpose_combo.setCurrentIndex(session.get('transpose_idx', 0))
        self.chord_window_combo.setCurrentIndex(session.get('chord_window_idx', 0))
        self.use_chords_cb.setChecked(session.get('use_chords', False))

        self._pr_tracks = session.get('tracks', [])
        self._pr_notes = notes

        # Restore piano roll state
        pr = self.piano_roll
        pr._gw2_only = session.get('gw2_only', False)
        pr._pitch_min = session.get('pitch_min', 21)
        pr._pitch_max = session.get('pitch_max', 108)
        pr.GW2_PITCH_MIN = session.get('gw2_pitch_min', 48)
        pr.GW2_PITCH_MAX = session.get('gw2_pitch_max', 83)
        pr._total_ms = session.get('total_ms', 0)
        pr._px_per_ms = session.get('px_per_ms', 0.2)

        bpm = 120
        pr.setNotes(self._pr_notes, self._pr_tracks, bpm)

        pr._scroll_x = session.get('scroll_x', 0)
        pr._scroll_y = session.get('scroll_y', 0)
        pr._total_ms = session.get('total_ms', 0)
        pr._update_scrollbars()
        pr.update()


        self._populate_track_list()
        self.piano_roll.updateSimplifiedNotes()
        self.piano_roll.update()
        self.play_btn.setEnabled(True)
        self.play_here_btn.setEnabled(True)
        self._enable_export(ahk=True, xml=(self.file_type == 'midi'), submit=True)

        label = os.path.basename(self.midi_path) if self.midi_path else "Restored session"
        self.setWindowTitle(f"Serenade Music Converter v{__version__} — {label}")
        simplified_count = sum(1 for n in self._pr_notes if n.simplified)
        msg = f"Restored previous session: {len(self._pr_notes)} notes"
        if simplified_count:
            msg += f" ({simplified_count} simplified)"
        self.log(msg)

    def closeEvent(self, event):
        """Save session before closing."""
        self.settings.setValue('window_geometry', self.saveGeometry())
        if self._pr_notes:
            self._save_session()
        super().closeEvent(event)


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = MainWindow()
    # If a file path is passed on the command line, load it instead of restoring session
    args = [a for a in sys.argv[1:] if not a.startswith('-')]
    if args and os.path.isfile(args[0]):
        window.show()
        window._load_file(args[0])
    else:
        window._restore_session()
        window.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
