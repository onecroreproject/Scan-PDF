import os
import sys
import django
import json

# Set up Django environment
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '../../')))
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'core.settings')
django.setup()

from video_downloader import services

def test_instagram():
    url = "https://www.instagram.com/reel/Cw9K5u-u3m_/" # Sample public Instagram reel URL
    print(f"Testing analyze for: {url}")
    try:
        result = services.analyze_video(url)
        print("TOTAL FORMATS:", len(result.get('formats', [])))
        for fmt in result.get('formats', []):
            if fmt.get('type') == 'Audio Only':
                print("FOUND AUDIO FORMAT:", json.dumps(fmt, indent=2))
                return
        print("NO AUDIO FORMAT FOUND IN RESPONSE!")
        print(json.dumps(result, indent=2))
    except Exception as e:
        print("ERROR:", e)

if __name__ == '__main__':
    test_instagram()
