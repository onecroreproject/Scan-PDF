import os
import subprocess
import tempfile
import uuid
import shutil
from pathlib import Path
from .utils import get_output_path

# Supported video formats
SUPPORTED_FORMATS = ['mp4', 'avi', 'mov', 'mkv', 'webm', 'wmv', 'flv', '3gp']

def convert_video_format(input_path, original_filename, output_format):
    """
    Converts a video file from one format to another using FFmpeg.
    
    Args:
        input_path (str): The absolute path to the uploaded temporary video file.
        original_filename (str): The original filename of the uploaded file.
        output_format (str): The desired output format (e.g., 'mp4', 'avi').
        
    Returns:
        str: The absolute path to the converted temporary video file.
        
    Raises:
        ValueError: If the output format is not supported or the input file doesn't exist.
        RuntimeError: If FFmpeg fails to convert the video.
    """
    output_format = output_format.lower().strip()
    
    if output_format not in SUPPORTED_FORMATS:
        raise ValueError(f"Output format '{output_format}' is not supported.")
        
    if not os.path.exists(input_path):
        raise ValueError("Input file does not exist.")

    # Create a safe output path using the project's utility
    output_path = get_output_path(original_filename, output_format, suffix='_converted')
    
    # Construct FFmpeg command
    command = [
        'ffmpeg',
        '-i', str(input_path),
        '-y',
    ]

    # Add specific encoding parameters based on the format to prevent codec/resolution errors
    if output_format == '3gp':
        # 3GP standard H.263 codec only supports exact resolutions (like 352x288 or 704x576). 
        # Using mpeg4 allows us to use any resolution, but we still scale it down to 
        # 720p max to ensure compatibility with devices that use 3GP.
        command.extend([
            '-vf', r'scale=-2:min(ih\,720)',
            '-c:v', 'mpeg4',
            '-c:a', 'aac',
            '-ar', '8000', # Standard audio rate for 3GP
        ])
    else:
        # For modern formats, use ultrafast preset to prevent network timeouts
        # and ensure smooth conversion speeds for 4K/60fps videos.
        command.extend([
            '-preset', 'ultrafast',
        ])
        
    # Finally, append the output path
    command.append(str(output_path))
    
    try:
        # Run FFmpeg securely using subprocess
        # capture_output to prevent terminal spam and to get error messages if it fails
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            check=True
        )
    except subprocess.CalledProcessError as e:
        # If it fails, make sure to clean up the empty/partial output file
        if os.path.exists(output_path):
            try:
                os.remove(output_path)
            except OSError:
                pass
        
        # Log or expose the error message to help with debugging
        error_msg = e.stderr if e.stderr else str(e)
        raise RuntimeError(f"FFmpeg conversion failed: {error_msg}")
        
    except FileNotFoundError:
        # This happens if ffmpeg is not installed or not in PATH
        if os.path.exists(output_path):
            try:
                os.remove(output_path)
            except OSError:
                pass
        raise RuntimeError("FFmpeg is not installed or not available in the system PATH.")
        
    return output_path


def convert_images_to_video(image_paths, durations, audio_path=None, audio_start=None, audio_end=None):
    """
    Creates an MP4 video from a sequence of images and durations, with optional audio.
    """
    output_path = get_output_path("merged_video", "mp4")
    
    with tempfile.TemporaryDirectory() as work_dir:
        work_dir_path = Path(work_dir)
        clip_paths = []
        
        for index, (image_path, duration) in enumerate(zip(image_paths, durations)):
            clip_path = work_dir_path / f"clip_{index}.mp4"
            command = [
                "ffmpeg", "-y",
                "-loop", "1",
                "-i", str(image_path),
                "-t", str(duration),
                "-vf", (
                    "scale=1280:720:"
                    "force_original_aspect_ratio=decrease,"
                    "pad=1280:720:"
                    "(ow-iw)/2:"
                    "(oh-ih)/2:"
                    "color=black"
                ),
                "-r", "30",
                "-c:v", "libx264",
                "-pix_fmt", "yuv420p",
                "-an",
                "-movflags", "+faststart",
                str(clip_path)
            ]
            try:
                subprocess.run(command, capture_output=True, text=True, check=True)
            except subprocess.CalledProcessError as e:
                raise RuntimeError(f"FFmpeg image to clip failed: {e.stderr}")
            except FileNotFoundError:
                raise RuntimeError("FFmpeg is not installed.")
                
            clip_paths.append(clip_path)
            
        video_only = work_dir_path / "video_only.mp4"
        concat_file = work_dir_path / "concat.txt"
        
        with open(concat_file, "w", encoding="utf-8") as f:
            for clip in clip_paths:
                f.write(f"file '{clip.resolve().as_posix()}'\n")
                
        concat_cmd = [
            "ffmpeg", "-y",
            "-f", "concat",
            "-safe", "0",
            "-i", str(concat_file),
            "-c", "copy",
            str(video_only)
        ]
        
        try:
            subprocess.run(concat_cmd, capture_output=True, text=True, check=True)
        except subprocess.CalledProcessError as e:
            raise RuntimeError(f"FFmpeg concat failed: {e.stderr}")
        
        if audio_path:
            audio_cmd = ["ffmpeg", "-y"]
            
            # 1. Video input
            audio_cmd.extend(["-i", str(video_only)])
            
            # 2. Audio input (with optional trim)
            if audio_start is not None and audio_end is not None:
                audio_cmd.extend(["-ss", str(audio_start), "-to", str(audio_end)])
            
            audio_cmd.extend(["-i", str(audio_path)])
            
            # 3. Muxing and encoding
            audio_cmd.extend([
                "-map", "0:v:0",
                "-map", "1:a:0",
                "-c:v", "copy",
                "-c:a", "aac",
                "-b:a", "192k",
                "-shortest",
                "-movflags", "+faststart",
                str(output_path)
            ])
            try:
                subprocess.run(audio_cmd, capture_output=True, text=True, check=True)
            except subprocess.CalledProcessError as e:
                raise RuntimeError(f"FFmpeg audio mux failed: {e.stderr}")
        else:
            # If no audio, still apply faststart to the concatenated video
            shutil.copy2(video_only, output_path)
            subprocess.run(["ffmpeg", "-y", "-i", str(output_path), "-c", "copy", "-movflags", "+faststart", str(output_path) + "_fs.mp4"], capture_output=True, check=True)
            shutil.move(str(output_path) + "_fs.mp4", output_path)

    # Validation step using FFprobe
    if not os.path.exists(output_path) or os.path.getsize(output_path) == 0:
        raise RuntimeError("FFmpeg generated an empty or missing file.")
        
    try:
        probe_cmd = [
            "ffprobe", "-v", "error", "-select_streams", "v:0", 
            "-show_entries", "stream=codec_type,width,height", 
            "-of", "csv=p=0", str(output_path)
        ]
        probe = subprocess.run(probe_cmd, capture_output=True, text=True, check=True).stdout.strip()
        if not probe or "video" not in probe:
            raise RuntimeError("Generated file has no valid video stream.")
    except Exception as e:
        raise RuntimeError(f"FFprobe validation failed: {str(e)}")
        
    return str(output_path)
            
    return str(output_path)

