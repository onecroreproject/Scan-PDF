# ScanPDF / All-In-One PDF & Media Tools

A comprehensive Django-based web application offering a suite of tools for processing PDFs, images, audio, video, dynamic QR codes, and short URLs.

## Features & Core Modules

- **Converter (`converter`)**: Tools for document conversion (e.g., PDF processing).
- **Image Processor (`image_processor`)**: Utilities for image manipulation and conversion.
- **Audio & Video Tools (`audio_processor`, `audio_replacement`, `video_downloader`)**: Audio editing, extracting, video downloading, and replacement workflows.
- **Dynamic QR & Short URLs (`dynamic_qr`)**: Create, manage, and track Dynamic QR Codes and Shortened URLs with rich analytics.
- **Services & Billing (`services`)**: Handles subscription plans, pricing tiers, payment processing, and usage limits.
- **Custom Admin (`custom_admin`)**: A custom-built dashboard for system management, user administration, subscriptions, and platform insights.

## Tech Stack

- **Backend**: Django (Python)
- **Database**: SQLite (default for development) 
- **Media Processing**: FFmpeg, MoviePy, Pydub

---

## Local Development Setup

### 1. Requirements

- Python 3.10+
- Virtual Environment

### 2. Installation

```bash
# Create and activate a virtual environment
python -m venv .venv

# Windows
.venv\Scripts\activate
# Linux/Mac
source .venv/bin/activate

# Install dependencies
pip install -r requirements.txt
```

### 3. Database Initialization & Seeding

```bash
# Run migrations
python manage.py migrate

# Create an admin user
python manage.py createsuperuser

# (Optional) Seed the Plans & Pricing features for Short URLs
python manage.py seed_shorturl_plan_features
```

### 4. Running the Server

```bash
python manage.py runserver
```

---

## FFmpeg Configuration (Media Processing)

This project’s audio/video tools (Pydub + MoviePy) are wired to **prefer a project-bundled FFmpeg** so you don’t need a global/system install on Windows.

### Folder Layout

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

### How it resolves FFmpeg

Resolution order is:
1. `ffmpeg/bin` inside the project (preferred)
2. Environment variables: `FFMPEG_BINARY`, `FFPROBE_BINARY`, `IMAGEIO_FFMPEG_EXE`
3. System `PATH` (`ffmpeg`, `ffprobe`)
4. `imageio-ffmpeg` managed binary (ffmpeg only; ffprobe is optional)

### Verify FFmpeg works

```bash
python manage.py check_ffmpeg
```

### Production / Linux notes

- If you **bundle Linux binaries** into `ffmpeg/bin/`, the app will use them.
- If you **don’t bundle**, the app can still work via `ffmpeg` from system `PATH`, and/or `imageio-ffmpeg` (already in `requirements.txt`).
- This repo also lists `ffmpeg` in `Aptfile` and `packages.txt` as a **Linux fallback** for buildpack-style environments.
