# Piano Roll Editing

The piano roll is the main editing area where you interact with notes visually. Pitch is on the vertical axis (piano keys on the left), time on the horizontal axis (measure numbers at the top).

## Navigation

- **Scroll vertically** — mouse wheel
- **Scroll horizontally** — Shift + mouse wheel
- **Zoom in/out** — Ctrl + mouse wheel, or Ctrl+= / Ctrl+-
- **Home** / **End** — scroll to the start or end of the song
- The **minimap** at the bottom shows an overview of all notes; click to jump to a position

## Selecting Notes

- **Click** a note to select it (deselects others)
- **Ctrl+Click** to add/remove a note from the selection
- **Click empty space** and drag to box-select multiple notes
- **Ctrl+A** to select all visible notes
- **Escape** to deselect all notes and clear any range selection
- Selected notes are highlighted with a bright outline

## Moving Notes

- **Drag** a selected note to move it (time and pitch)
- All selected notes move together
- Notes snap to the grid (based on BPM and zoom level)

## Resizing Notes

- Hover near the **right edge** of a note — the cursor changes to a resize handle
- **Drag** to change the note's duration

## Drawing Notes

- Click the **✏ Draw** button (top-right) or press **Ctrl+D** to enter draw mode
- **Click** on the piano roll to place a new note
- **Click and drag** horizontally to set the note's duration while drawing
- Click **⬚ Select** or press Ctrl+D again to return to select mode

## Context Menu

- **Right-click** a note to open the context menu with options like delete, octave shift, etc.
- If no notes are currently selected, right-clicking a note **automatically selects it** first, so the context menu operates on the clicked note

## Deleting Notes

- Select notes and press **Del** to delete them
- Or right-click a note and choose Delete

## Copy & Paste

- **Ctrl+C** to copy selected notes
- **Ctrl+V** to paste — notes are placed at the current scroll position

## Undo & Redo

- **Ctrl+Z** to undo the last action
- **Ctrl+Y** to redo

## Range Selection

- **Click and drag** on the **ruler** (measure bar at the top) to select a time range
- The range is shown as a highlighted region
- **Right-click** the ruler for trim options (trim to range, delete outside range)
- When a range is active, playback and export operate on that range only

## GW2 Pitch Range

The red dashed lines on the piano roll show the playable range for the selected GW2 instrument. Notes outside this range will need octave adjustment (via track tools) or will be flagged as out-of-range in the status bar.
