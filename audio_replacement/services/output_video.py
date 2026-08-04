import os
import shutil
import uuid
from django.conf import settings


class OutputVideo:

    def __init__(self, video_path):

        self.video_path = video_path

    # ---------------------------------------
    # Create Output Directory
    # ---------------------------------------
    def create_directory(self):

        output_dir = os.path.join(
            settings.MEDIA_ROOT,
            "output"
        )

        os.makedirs(output_dir, exist_ok=True)

        return output_dir

    # ---------------------------------------
    # Generate Unique File Name
    # ---------------------------------------
    def generate_filename(self):

        return f"{uuid.uuid4().hex}.mp4"

    # ---------------------------------------
    # Save Output Video
    # ---------------------------------------
    def save(self):

        output_dir = self.create_directory()

        filename = self.generate_filename()

        destination = os.path.join(
            output_dir,
            filename
        )

        shutil.copy2(
            self.video_path,
            destination
        )

        return destination

    # ---------------------------------------
    # Get Relative Media Path
    # ---------------------------------------
    def get_media_path(self):

        saved_file = self.save()

        return os.path.relpath(
            saved_file,
            settings.MEDIA_ROOT
        )

    # ---------------------------------------
    # Delete Temporary File
    # ---------------------------------------
    def delete_temp(self):

        if os.path.exists(self.video_path):

            os.remove(self.video_path)

            return True

        return False

    # ---------------------------------------
    # Save and Clean
    # ---------------------------------------
    def finalize(self):

        media_path = self.get_media_path()

        self.delete_temp()

        return media_path