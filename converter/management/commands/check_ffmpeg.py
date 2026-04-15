import subprocess

from django.core.management.base import BaseCommand

from converter.media_binaries import ensure_ffmpeg_configured


class Command(BaseCommand):
    help = "Verify FFmpeg/FFprobe discovery and executability."

    def handle(self, *args, **options):
        ffmpeg_path, ffprobe_path = ensure_ffmpeg_configured()

        self.stdout.write(f"ffmpeg:  {ffmpeg_path or '<not found>'}")
        self.stdout.write(f"ffprobe: {ffprobe_path or '<not found>'}")

        if not ffmpeg_path:
            raise SystemExit(
                "FFmpeg was not found. Bundle it in ffmpeg/bin, set FFMPEG_BINARY, "
                "or ensure 'ffmpeg' is on PATH."
            )

        try:
            completed = subprocess.run(
                [ffmpeg_path, "-version"],
                check=True,
                capture_output=True,
                text=True,
            )
        except FileNotFoundError as exc:
            raise SystemExit(f"FFmpeg path is not executable: {ffmpeg_path}") from exc
        except subprocess.CalledProcessError as exc:
            stderr = (exc.stderr or "").strip()
            raise SystemExit(f"FFmpeg ran but returned error.\n\n{stderr}") from exc

        first_line = (completed.stdout or "").splitlines()[0] if completed.stdout else ""
        self.stdout.write(first_line or "FFmpeg OK")

