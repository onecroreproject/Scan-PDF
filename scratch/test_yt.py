import sys
import os
import json
import traceback

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
import django
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'allinone.settings')
django.setup()

from video_downloader.services import analyze_video

def main():
    urls = [
        "https://www.youtube.com/watch?v=dQw4w9WgXcQ",
        "https://youtube.com/shorts/3Ozvq_gVbF8"
    ]
    for url in urls:
        print(f"Testing URL: {url}")
        try:
            res = analyze_video(url)
            print("SUCCESS! Formats count:", len(res.get('formats', [])))
        except Exception as e:
            print("ERROR:")
            print(traceback.format_exc())
            if hasattr(e, 'raw_error'):
                print("RAW ERROR:")
                print(e.raw_error)

if __name__ == '__main__':
    main()
