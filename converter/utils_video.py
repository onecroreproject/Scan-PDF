import os
import subprocess
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
