import sys
sys.path.append(r'r:\DLK-Project\Scan-PDF')

import logging
logging.basicConfig(level=logging.DEBUG)

from media_tools.services.downloader_service import extract_threads_images
try:
    res = extract_threads_images('https://www.threads.com/share/DeHKj4hgd3m/')
    print(res)
except Exception as e:
    print("FAILED:", e)
