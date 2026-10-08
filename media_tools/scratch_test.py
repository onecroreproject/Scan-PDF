import sys
import os
import django

# Setup Django environment
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
os.environ.setdefault("DJANGO_SETTINGS_MODULE", "allinone.settings")
django.setup()

from media_tools.services.downloader_service import extract_instagram_media
import urllib.request

url = "https://www.instagram.com/p/DeLbDOzqnxN/"
print(f"Testing URL: {url}")

try:
    res = extract_instagram_media(url)
    print("SUCCESS!")
    print(res)
except Exception as e:
    print(f"FAILED: {e}")
    import traceback
    traceback.print_exc()

# Let's also print what the embed URL returns
print("\n--- Embed HTML ---")
embed_url = "https://www.instagram.com/p/DeLbDOzqnxN/embed/captioned/"
req = urllib.request.Request(embed_url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
try:
    with urllib.request.urlopen(req, timeout=10) as resp:
        html = resp.read().decode('utf-8')
        print(f"Embed HTML Length: {len(html)}")
        import re
        script_matches = re.findall(r'<script[^>]*>([\s\S]*?)</script>', html)
        print(f"Script tags: {len(script_matches)}")
        for s in script_matches:
            if 'shortcode' in s or 'image_versions2' in s or 'DeLbDOzqnxN' in s:
                print("FOUND in script:", len(s), "bytes")
except Exception as e:
    print("Embed failed:", e)
