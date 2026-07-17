from .views import TOOLS
from image_processor.views import IMAGE_TOOLS
from django.urls import reverse

def tools_processor(request):
    """Make all tools available to all templates, grouped by category."""
    grouped_tools = {}
    
    # Combined dictionary for search metadata
    all_combined = {**TOOLS, **IMAGE_TOOLS}

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
    }

    CATEGORY_ORDER = [
        'convert', 'pdf-tools', 'image-tools', 'image-pro', 'image-conv',
        'generate', 'ai-tools', 'other', 'audio-tools'
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

    # Add Video Downloader tools to search
    metadata['video-downloader-universal'] = {
        'title': 'Universal Video Downloader',
        'icon': 'download-cloud',
        'description': 'Download videos from YouTube, Facebook, Instagram, TikTok, and more.',
        'slug': 'video-downloader',
        'url': reverse('video_downloader:index')
    }
    metadata['youtube-downloader'] = {
        'title': 'YouTube Video Downloader',
        'icon': 'youtube',
        'description': 'Download YouTube videos easily in MP4 or MP3 format.',
        'slug': 'youtube-downloader',
        'url': reverse('video_downloader:youtube_downloader')
    }
    metadata['facebook-downloader'] = {
        'title': 'Facebook Video Downloader',
        'icon': 'facebook',
        'description': 'Download Facebook videos directly to your device.',
        'slug': 'facebook-downloader',
        'url': reverse('video_downloader:facebook_downloader')
    }
    metadata['twitter-downloader'] = {
        'title': 'X (Twitter) Video Downloader',
        'icon': 'twitter',
        'description': 'Download videos and GIFs from X (formerly Twitter).',
        'slug': 'twitter-downloader',
        'url': reverse('video_downloader:twitter_downloader')
    }
    metadata['instagram-downloader'] = {
        'title': 'Instagram Video Downloader',
        'icon': 'instagram',
        'description': 'Download Instagram Reels, IGTV, and videos.',
        'slug': 'instagram-downloader',
        'url': reverse('video_downloader:instagram_downloader')
    }
    metadata['tiktok-downloader'] = {
        'title': 'TikTok Video Downloader',
        'icon': 'music-2',
        'description': 'Download TikTok videos without watermark.',
        'slug': 'tiktok-downloader',
        'url': reverse('video_downloader:tiktok_downloader')
    }
    metadata['vimeo-downloader'] = {
        'title': 'Vimeo Video Downloader',
        'icon': 'video',
        'description': 'Download Vimeo videos in HD quality.',
        'slug': 'vimeo-downloader',
        'url': reverse('video_downloader:vimeo_downloader')
    }
    metadata['reddit-downloader'] = {
        'title': 'Reddit Video Downloader',
        'icon': 'hash',
        'description': 'Download Reddit videos with audio.',
        'slug': 'reddit-downloader',
        'url': reverse('video_downloader:reddit_downloader')
    }
    metadata['dailymotion-downloader'] = {
        'title': 'Dailymotion Video Downloader',
        'icon': 'play-circle',
        'description': 'Download Dailymotion videos in high quality.',
        'slug': 'dailymotion-downloader',
        'url': reverse('video_downloader:dailymotion_downloader')
    }


    return {
        'grouped_tools': ordered,
        'all_tools_metadata': metadata,
    }
