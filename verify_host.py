import os
import sys
import subprocess
import urllib.request
import platform

print("="*50)
print("  ScanPDF Hostinger Diagnostic & Repair Tool")
print("="*50)

# 1. Check Python Version
print("\n[1] Checking Python Environment...")
print(f"Python Version: {sys.version.split(' ')[0]}")
if sys.version_info < (3, 8):
    print("WARNING: Python 3.8+ is recommended for yt-dlp.")
else:
    print("Python version is compatible.")

# 2. Update yt-dlp
print("\n[2] Updating yt-dlp to the latest version...")
try:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "-U", "yt-dlp"])
    print("yt-dlp updated successfully.")
except subprocess.CalledProcessError:
    print("FAILED to update yt-dlp. Check pip permissions.")

# 3. Check FFmpeg
print("\n[3] Checking FFmpeg installation...")
def check_binary(name):
    try:
        if platform.system() == "Windows":
            cmd = ["where", name]
        else:
            cmd = ["which", name]
        result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        if result.returncode == 0:
            return result.stdout.strip()
    except Exception:
        pass
    return None

ffmpeg_path = check_binary("ffmpeg")
ffprobe_path = check_binary("ffprobe")

if ffmpeg_path:
    print(f"FFmpeg found at: {ffmpeg_path}")
else:
    print("WARNING: FFmpeg is NOT installed globally or not in PATH!")
    print("Hostinger Shared Hosting usually does not provide FFmpeg. Video merging will fail.")
    print("Solution: Switch to a VPS, or download statically compiled FFmpeg and set FFMPEG_BIN_DIR in settings.py.")

# 4. Check Network Restrictions
print("\n[4] Checking Outbound Network Restrictions...")
urls_to_test = {
    "YouTube": "https://www.youtube.com",
    "Instagram": "https://www.instagram.com/accounts/login/",
    "Reddit": "https://www.reddit.com"
}

for site, url in urls_to_test.items():
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        response = urllib.request.urlopen(req, timeout=10)
        print(f"✅ {site} is accessible (HTTP {response.getcode()})")
    except urllib.error.HTTPError as e:
        print(f"❌ {site} blocked access: HTTP {e.code} (Your server IP is likely flagged)")
    except Exception as e:
        print(f"❌ {site} connection failed: {e}")

# 5. Check Permissions
print("\n[5] Checking File Permissions...")
base_dir = os.path.dirname(os.path.abspath(__file__))
scratch_dir = os.path.join(base_dir, 'scratch')
media_dir = os.path.join(base_dir, 'media', 'video_downloads')

for d in [scratch_dir, media_dir]:
    os.makedirs(d, exist_ok=True)
    if os.access(d, os.W_OK):
        print(f"✅ Write permission OK for: {d}")
    else:
        print(f"❌ WRITE PERMISSION DENIED for: {d}")

print("\n" + "="*50)
print("Diagnostic Complete.")
print("If Instagram/YouTube returned HTTP 429 or 403, your Hostinger IP is banned.")
print("You MUST upload a valid 'cookies.txt' from your browser to the root directory to bypass IP bans, OR migrate to a VPS.")
