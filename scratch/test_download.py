import os
import sys
import django

# Setup django
sys.path.append(r'r:\DLK-Project\Scan-PDF')
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'allinone.settings')
django.setup()

from video_downloader.services import download_format

try:
    url = 'https://www.facebook.com/watch/?v=1328906647986064'
    filepath, title = download_format(url, 'bestaudio/best', 'Audio Only')
    print(f"SUCCESS: {filepath} ({title})")
except Exception as e:
    import traceback
    print("FAILED!")
    traceback.print_exc()
