from django.apps import AppConfig

class VideoProcessorConfig(AppConfig):
    default_auto_field = 'django.db.models.BigAutoField'
    name = 'video_processor'

    def ready(self):
        # Start background temp file cleanup scheduler
        try:
            from .cleanup import CleanupScheduler
            from django.conf import settings
            from .utils import get_uploads_dir, get_outputs_dir, get_chunks_dir
            dirs = [get_uploads_dir(), get_outputs_dir(), get_chunks_dir()]
            scheduler = CleanupScheduler(dirs, interval=300, max_age=getattr(settings, 'VIDEO_TEMP_MAX_AGE', 600))
            scheduler.start()
        except Exception:
            import logging
            logging.getLogger(__name__).warning("Could not start video cleanup scheduler", exc_info=True)
