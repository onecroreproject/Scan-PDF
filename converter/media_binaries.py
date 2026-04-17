import os
import shutil
from functools import lru_cache
from pathlib import Path


def _binary_name(base: str) -> str:
    return f"{base}.exe" if os.name == "nt" else base


def _get_winget_ffmpeg_bin_dir() -> str | None:
    """
    Windows-only convenience: if the user installed FFmpeg via WinGet
    (Gyan.FFmpeg package), discover its bin directory.
    """
    local_app_data = os.environ.get("LOCALAPPDATA", "")
    if not local_app_data:
        return None
    pattern_root = Path(local_app_data) / "Microsoft" / "WinGet" / "Packages"
    if not pattern_root.exists():
        return None

    # Avoid globbing huge trees unnecessarily; keep it tight to the known package prefix.
    pkg_root = pattern_root / "Gyan.FFmpeg_Microsoft.Winget.Source_8wekyb3d8bbwe"
    if not pkg_root.exists():
        return None

    # Expected: ffmpeg-*-full_build/bin
    candidates: list[Path] = []
    try:
        for child in pkg_root.iterdir():
            if child.is_dir() and child.name.startswith("ffmpeg-") and child.name.endswith("_full_build"):
                bin_dir = child / "bin"
                if bin_dir.exists():
                    candidates.append(bin_dir)
    except OSError:
        return None

    if not candidates:
        return None
    candidates.sort(reverse=True)
    return str(candidates[0])


def _try_get_django_base_dir() -> Path | None:
    try:
        from django.conf import settings  # type: ignore
    except Exception:
        return None

    try:
        base_dir = getattr(settings, "BASE_DIR", None)
    except Exception:
        return None

    if not base_dir:
        return None
    return Path(base_dir)


def _project_root_fallback() -> Path:
    # converter/media_binaries.py -> converter -> project root
    return Path(__file__).resolve().parents[1]


def resolve_ffmpeg_paths(base_dir: Path | None = None) -> tuple[str | None, str | None]:
    """
    Resolve FFmpeg/FFprobe paths with this preference order:
    1) Project-bundled ffmpeg/bin
    2) Environment variables (FFMPEG_BINARY / FFPROBE_BINARY / IMAGEIO_FFMPEG_EXE)
    3) System PATH (shutil.which)
    4) imageio-ffmpeg managed binary (ffmpeg only), with local ffprobe neighbor if present
    5) Windows WinGet-installed FFmpeg (convenience)
    """
    base_dir = base_dir or _try_get_django_base_dir() or _project_root_fallback()

    candidates_ffmpeg: list[str] = []
    candidates_ffprobe: list[str] = []

    # 1) Project-bundled
    bin_dir = base_dir / "ffmpeg" / "bin"
    bundled_ffmpeg = bin_dir / _binary_name("ffmpeg")
    bundled_ffprobe = bin_dir / _binary_name("ffprobe")
    candidates_ffmpeg.append(str(bundled_ffmpeg))
    candidates_ffprobe.append(str(bundled_ffprobe))

    # 2) Environment variables
    env_ffmpeg = os.environ.get("FFMPEG_BINARY") or os.environ.get("IMAGEIO_FFMPEG_EXE")
    env_ffprobe = os.environ.get("FFPROBE_BINARY")
    if env_ffmpeg:
        candidates_ffmpeg.append(env_ffmpeg)
    if env_ffprobe:
        candidates_ffprobe.append(env_ffprobe)

    # 3) PATH
    which_ffmpeg = shutil.which("ffmpeg")
    which_ffprobe = shutil.which("ffprobe")
    if which_ffmpeg:
        candidates_ffmpeg.append(which_ffmpeg)
    if which_ffprobe:
        candidates_ffprobe.append(which_ffprobe)

    # 4) imageio-ffmpeg
    try:
        import imageio_ffmpeg  # type: ignore

        candidates_ffmpeg.append(imageio_ffmpeg.get_ffmpeg_exe())
    except Exception:
        pass

    # 5) WinGet (Windows only) — keep late; local bundling should win first
    if os.name == "nt":
        winget_bin = _get_winget_ffmpeg_bin_dir()
        if winget_bin:
            candidates_ffmpeg.append(str(Path(winget_bin) / _binary_name("ffmpeg")))
            candidates_ffprobe.append(str(Path(winget_bin) / _binary_name("ffprobe")))

    def _pick_existing(paths: list[str]) -> str | None:
        for p in paths:
            if p and os.path.exists(p):
                return p
        return None

    ffmpeg_path = _pick_existing(candidates_ffmpeg)
    ffprobe_path = _pick_existing(candidates_ffprobe)

    # If ffprobe not found, try sibling next to ffmpeg.
    if ffmpeg_path and not ffprobe_path:
        sibling = str(Path(ffmpeg_path).parent / _binary_name("ffprobe"))
        if os.path.exists(sibling):
            ffprobe_path = sibling

    return ffmpeg_path, ffprobe_path


def configure_moviepy(ffmpeg_path: str | None) -> None:
    if not ffmpeg_path:
        return
    # MoviePy (via imageio-ffmpeg) respects IMAGEIO_FFMPEG_EXE.
    os.environ["IMAGEIO_FFMPEG_EXE"] = ffmpeg_path


@lru_cache(maxsize=1)
def ensure_ffmpeg_configured() -> tuple[str | None, str | None]:
    """
    Resolve FFmpeg paths once per process and configure MoviePy.
    Safe to call multiple times.
    """
    ffmpeg_path, ffprobe_path = resolve_ffmpeg_paths()
    configure_moviepy(ffmpeg_path)
    return ffmpeg_path, ffprobe_path

