import subprocess
import json
import sys
import os

def validate_media(filepath):
    if not os.path.exists(filepath):
        print(f"File not found: {filepath}")
        return False
        
    if os.path.getsize(filepath) == 0:
        print(f"File is 0 bytes: {filepath}")
        return False

    cmd = [
        'ffprobe', 
        '-v', 'quiet', 
        '-print_format', 'json', 
        '-show_format', 
        '-show_streams', 
        filepath
    ]
    
    try:
        result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        data = json.loads(result.stdout)
        
        streams = data.get('streams', [])
        format_info = data.get('format', {})
        
        video_streams = [s for s in streams if s.get('codec_type') == 'video']
        audio_streams = [s for s in streams if s.get('codec_type') == 'audio']
        
        print(f"--- VALIDATION REPORT: {os.path.basename(filepath)} ---")
        print(f"Size: {format_info.get('size')} bytes")
        print(f"Duration: {format_info.get('duration')} seconds")
        print(f"Format: {format_info.get('format_name')}")
        print(f"Video Streams: {len(video_streams)}")
        print(f"Audio Streams: {len(audio_streams)}")
        
        if video_streams:
            print(f"Video Codec: {video_streams[0].get('codec_name')}")
        if audio_streams:
            print(f"Audio Codec: {audio_streams[0].get('codec_name')}")
            
        print("-------------------------------------------------")
        return True
    except Exception as e:
        print(f"FFprobe failed for {filepath}: {e}")
        return False

if __name__ == '__main__':
    if len(sys.argv) > 1:
        validate_media(sys.argv[1])
    else:
        print("Please provide a file path.")
