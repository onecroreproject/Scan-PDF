from .views import TOOLS
from image_processor.views import IMAGE_TOOLS
from django.urls import reverse

# Video tools metadata for search integration
VIDEO_TOOLS = {
    'video-converter': {'title': 'Video Converter', 'icon': 'refresh-cw', 'description': 'Convert videos between MP4, AVI, MOV, MKV, and WEBM.', 'category': 'video-tools'},
    'image-to-video': {'title': 'Image to Video', 'icon': 'image', 'description': 'Create slideshow videos from your photos.', 'category': 'video-tools'},
    'video-editor': {'title': 'Video Editor', 'icon': 'scissors', 'description': 'Trim, cut, rotate, resize, crop, and edit videos.', 'category': 'video-tools'},
    'video-compressor': {'title': 'Video Compressor', 'icon': 'archive', 'description': 'Reduce video file size while maintaining quality.', 'category': 'video-tools'},
    'video-merger': {'title': 'Video Merger', 'icon': 'combine', 'description': 'Combine multiple videos into one file.', 'category': 'video-tools'},
    'video-trimmer': {'title': 'Video Trimmer', 'icon': 'crop', 'description': 'Trim videos to specific start and end times.', 'category': 'video-tools'},
    'gif-maker': {'title': 'GIF Maker', 'icon': 'image', 'description': 'Convert video clips into animated GIFs.', 'category': 'video-tools'},
    'audio-extractor': {'title': 'Audio Extractor', 'icon': 'music', 'description': 'Extract audio from videos as MP3, WAV, or AAC.', 'category': 'video-tools'},
    'watermark-tool': {'title': 'Watermark Tool', 'icon': 'stamp', 'description': 'Add text or image watermarks to videos.', 'category': 'video-tools'},
    'subtitle-overlay': {'title': 'Subtitle Overlay', 'icon': 'subtitles', 'description': 'Burn subtitles permanently into videos.', 'category': 'video-tools'},
}

VIDEO_TOOL_URLS = {
    'video-converter': 'video_processor:converter',
    'image-to-video': 'video_processor:image_to_video',
    'video-editor': 'video_processor:video_editor',
    'video-compressor': 'video_processor:compressor',
    'video-merger': 'video_processor:merger',
    'video-trimmer': 'video_processor:trimmer',
    'gif-maker': 'video_processor:gif_maker',
    'audio-extractor': 'video_processor:audio_extractor',
    'watermark-tool': 'video_processor:watermark',
    'subtitle-overlay': 'video_processor:subtitle_overlay',
}

def tools_processor(request):
    """Make all tools available to all templates, grouped by category."""
    grouped_tools = {}
    
    # Combined dictionary for search metadata
    all_combined = {**TOOLS, **IMAGE_TOOLS, **VIDEO_TOOLS}

    # Category display names
    CATEGORY_LABELS = {
        'convert': 'Convert to/from PDF',
        'pdf-tools': 'PDF Tools',
        'image-tools': 'Image Tools',
        'image-pro': 'Image Tools',
        'image-conv': 'Image Tools',
        'generate': 'Smart Creators',
        'ai-tools': 'AI Generation',
        'other': 'Utilities',
        'audio-tools': 'Audio Editor',
        'video-tools': 'Video Tools',
    }

    CATEGORY_ORDER = [
        'convert', 'pdf-tools', 'image-tools', 'image-pro', 'image-conv',
        'generate', 'ai-tools', 'video-tools', 'other', 'audio-tools'
    ]

    for slug, data in all_combined.items():
        # Skip if coming soon and marked as such in the source
        if data.get('is_coming_soon') and slug in IMAGE_TOOLS:
            continue
            
        cat = data.get('category', 'other')
        if cat not in grouped_tools:
            grouped_tools[cat] = {
                'label': CATEGORY_LABELS.get(cat, cat.replace('-', ' ').title()),
                'tools': []
            }

        def _app_name(s):
            if slug in IMAGE_TOOLS: return 'image_processor'
            if slug in VIDEO_TOOLS: return 'video_processor'
            return 'converter'
        grouped_tools[cat]['tools'].append({
            'title': data.get('title'),
            'icon': data.get('icon'),
            'slug': slug,
            'is_coming_soon': data.get('is_coming_soon', False),
            'app_name': _app_name(slug)
        })

    # Re-order the dict
    ordered = {}
    for cat in CATEGORY_ORDER:
        if cat in grouped_tools:
            ordered[cat] = grouped_tools[cat]
    for cat, info in grouped_tools.items():
        if cat not in ordered:
            ordered[cat] = info

    def _tool_url(s):
        if s in IMAGE_TOOLS:
            return reverse('image_processor:tool_page', args=[s])
        if s in VIDEO_TOOLS:
            return reverse(VIDEO_TOOL_URLS[s])
        return reverse('converter:convert_page', args=[s])

    # Prepare metadata for search
    metadata = {
        slug: {
            'title': data['title'],
            'icon': data['icon'],
            'description': data.get('description', ''),
            'slug': slug,
            'url': _tool_url(slug)
        }
        for slug, data in all_combined.items()
    }

    # Manually add Dynamic QR and Short URL to search
    is_dqr_user = request.session.get('is_dqr_user', False)
    
    metadata['dynamic-qr'] = {
        'title': 'Dynamic QR',
        'icon': 'qr-code',
        'description': 'Create and manage trackable dynamic QR codes with analytics.',
        'slug': 'dynamic-qr',
        'url': reverse('dynamic_qr:dashboard') if is_dqr_user else reverse('dynamic_qr:login')
    }
    metadata['short-url'] = {
        'title': 'Short URL',
        'icon': 'link',
        'description': 'Shorten URLs and track clicks with detailed analytics.',
        'slug': 'short-url',
        'url': reverse('dynamic_qr:short_url') if is_dqr_user else reverse('dynamic_qr:login')
    }

    return {
        'grouped_tools': ordered,
        'all_tools_metadata': metadata,
    }
