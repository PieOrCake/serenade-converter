# Building from Source

## Running Locally

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

The AppImage is built inside an Ubuntu 22.04 container using [Podman](https://podman.io/) (Docker also works with minor script edits).

```bash
./build-appimage.sh
```

This will:
1. Build the container image (first run only, cached afterwards)
2. Run PyInstaller inside the container to create a single-file executable
3. Assemble the AppDir structure with desktop integration files and Qt plugins
4. Package it as `Serenade_Music_Converter-x86_64.AppImage`

### Build Dependencies

The container handles all build dependencies automatically. On the host, you only need:
- **Podman** (or Docker)
- **bash**

### What's in the AppImage

| Component | Purpose |
|---|---|
| PyInstaller bundle | Python + PyQt6 + all dependencies in a single executable |
| AppRun | Entry point that sets up Qt plugin paths and environment |
| libqxdgdesktopportal.so | Qt platform theme plugin for native KDE/GNOME file dialogs |
| Desktop file + icons | Linux desktop integration (app name, icon, categories) |

### Rebuilding After Changes

Just run `./build-appimage.sh` again. The container image is cached, so subsequent builds only re-run PyInstaller and repackage.
