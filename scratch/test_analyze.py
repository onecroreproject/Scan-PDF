import os
import sys
import django
import json

# Set up Django environment
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '../../')))
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'core.settings')
django.setup()

from video_downloader import services

def test_facebook():
    url = "https://www.facebook.com/100085871358980/videos/969634731364539" # just a typical FB URL or the one in the example
    print(f"Testing analyze for: {url}")
    try:
        result = services.analyze_video(url)
        print("TOTAL FORMATS:", len(result.get('formats', [])))
        print(json.dumps(result, indent=2))
    except Exception as e:
        print("ERROR:", e)

if __name__ == '__main__':
    test_facebook()
