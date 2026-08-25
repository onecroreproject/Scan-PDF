from django.core.management.base import BaseCommand
import sys
import shutil
from video_downloader.services import _categorize_error
import yt_dlp

class Command(BaseCommand):
    help = 'Diagnostics for YouTube Downloader'

    def handle(self, *args, **options):
        self.stdout.write("--- Diagnostic Test for YouTube Downloader ---")
        
        # 1. yt-dlp version
        try:
            self.stdout.write(f"yt-dlp version: {yt_dlp.version.__version__} - PASS")
        except Exception as e:
            self.stdout.write(f"yt-dlp version: FAIL - {str(e)}")
            
        # 2. FFmpeg availability
        ffmpeg_path = shutil.which("ffmpeg")
        if ffmpeg_path:
            self.stdout.write(f"FFmpeg: PASS ({ffmpeg_path})")
        else:
            self.stdout.write("FFmpeg: FAIL (not found in PATH)")
            
        # 3. YouTube network test
        self.stdout.write("\nTesting YouTube network access...")
        test_url = "https://www.youtube.com/watch?v=jNQXAC9IVRw"
        
        try:
            opts = {
                'quiet': True,
                'no_warnings': True,
                'nocheckcertificate': True,
                'geo_bypass': True,
            }
            with yt_dlp.YoutubeDL(opts) as ydl:
                ydl.extract_info(test_url, download=False)
            self.stdout.write("YouTube network: PASS")
        except Exception as e:
            code, msg = _categorize_error(e, test_url)
            self.stdout.write("YouTube network: FAIL")
            self.stdout.write(f"Reason: {code}")
            self.stdout.write(f"Sanitized message: {msg}")
