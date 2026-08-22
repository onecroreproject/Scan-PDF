# All In One PDF — Media Processing (Local FFmpeg)

This project’s audio/video tools (Pydub + MoviePy) are wired to **prefer a project-bundled FFmpeg** so you don’t need a global/system install on Windows.

## Folder layout

Put binaries here:

- **Windows**
  - `ffmpeg/bin/ffmpeg.exe`
  - `ffmpeg/bin/ffprobe.exe`
- **Linux**
  - `ffmpeg/bin/ffmpeg`
  - `ffmpeg/bin/ffprobe`

The paths are exposed in Django settings:

- `FFMPEG_BIN_DIR`
- `FFMPEG_PATH`
- `FFPROBE_PATH`

## How it resolves FFmpeg

Resolution order is:

1. `ffmpeg/bin` inside the project (preferred)
2. Environment variables:
   - `FFMPEG_BINARY`, `FFPROBE_BINARY`
   - `IMAGEIO_FFMPEG_EXE` (MoviePy/ImageIO)
3. System `PATH` (`ffmpeg`, `ffprobe`)
4. `imageio-ffmpeg` managed binary (ffmpeg only; ffprobe is optional)

## Verify FFmpeg works

Run:

```bash
python manage.py check_ffmpeg
```

## Production / Linux notes

- If you **bundle Linux binaries** into `ffmpeg/bin/`, the app will use them.
- If you **don’t bundle**, the app can still work via:
  - `ffmpeg` from system `PATH`, and/or
  - `imageio-ffmpeg` (already in `requirements.txt`)
- This repo also lists `ffmpeg` in `Aptfile` and `packages.txt` as a **Linux fallback** for buildpack-style environments.

## HTML to Image Setup (Playwright)

This project uses Playwright for robust HTML to Image rendering. After running `pip install -r requirements.txt`, you must install the Chromium browser binary:

```bash
python -m playwright install chromium
```
