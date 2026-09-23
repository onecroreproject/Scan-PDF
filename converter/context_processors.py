from .views import TOOLS
from image_processor.views import IMAGE_TOOLS
from django.urls import reverse

def tools_processor(request):
    """Make all tools available to all templates, strictly grouped by category to avoid cross-contamination."""
    
    # Combined dictionary for search metadata
    all_combined = {**TOOLS, **IMAGE_TOOLS}

    # Category display names
    CATEGORY_LABELS = {
        'convert': 'Convert to/from PDF',
        'pdf-tools': 'PDF Tools',
        'pdf-edit': 'PDF Edit & Security',
        'pdf-adv': 'Advanced PDF',
        'image-tools': 'Image Editing',
        'image-pro': 'Image Editing',
        'image-conv': 'Image Converter',
        'generate': 'Smart Creators',
        'ai-tools': 'AI Generation',
        'other': 'Utilities',
        'audio-tools': 'Audio Editor',
    }

    # Desired order for PDF Tools Mega Menu
    PDF_CATEGORY_ORDER = ['convert', 'pdf-tools', 'pdf-edit', 'pdf-adv', 'other']
    
    # Desired order for Image Tools Mega Menu
    IMAGE_CATEGORY_ORDER = ['image-tools', 'image-pro', 'image-conv', 'generate', 'ai-tools']

    pdf_tools_grouped = {}
    for slug, data in TOOLS.items():
        cat = data.get('category', 'other')
        if cat not in pdf_tools_grouped:
            pdf_tools_grouped[cat] = {
                'label': CATEGORY_LABELS.get(cat, cat.replace('-', ' ').title()),
                'tools': []
            }
        pdf_tools_grouped[cat]['tools'].append({
            'title': data.get('title'),
            'icon': data.get('icon'),
            'slug': slug,
            'is_coming_soon': data.get('is_coming_soon', False),
            'app_name': 'converter'
        })

    image_tools_grouped = {}
    for slug, data in IMAGE_TOOLS.items():
        if data.get('is_coming_soon'):
            continue
            
        cat = data.get('category', 'other')
        if cat not in image_tools_grouped:
            image_tools_grouped[cat] = {
                'label': CATEGORY_LABELS.get(cat, cat.replace('-', ' ').title()),
                'tools': []
            }
        image_tools_grouped[cat]['tools'].append({
            'title': data.get('title'),
            'icon': data.get('icon'),
            'slug': slug,
            'is_coming_soon': data.get('is_coming_soon', False),
            'app_name': 'image_processor'
        })

    # Re-order the dicts
    ordered_pdf = {}
    for cat in PDF_CATEGORY_ORDER:
        if cat in pdf_tools_grouped:
            ordered_pdf[cat] = pdf_tools_grouped[cat]
    for cat, info in pdf_tools_grouped.items():
        if cat not in ordered_pdf:
            ordered_pdf[cat] = info

    ordered_image = {}
    for cat in IMAGE_CATEGORY_ORDER:
        if cat in image_tools_grouped:
            ordered_image[cat] = image_tools_grouped[cat]
    for cat, info in image_tools_grouped.items():
        if cat not in ordered_image:
            ordered_image[cat] = info

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

    # Video Tools and Link Tools for global navigation
    video_tools = [
        {'title': 'Converter', 'icon': 'video', 'url': reverse('converter:convert_page', args=['video-converter'])},
        {'title': 'Universal Downloader', 'icon': 'download-cloud', 'url': reverse('video_downloader:index')},
        {'title': 'YouTube', 'icon': 'youtube', 'url': reverse('video_downloader:youtube_downloader')},
        {'title': 'Trim Video', 'icon': 'scissors', 'url': reverse('media_tools:trim')},
        {'title': 'Merge Video', 'icon': 'combine', 'url': reverse('media_tools:merge')},
        {'title': 'Crop Video', 'icon': 'crop', 'url': reverse('media_tools:crop')},
        {'title': 'Resize Video', 'icon': 'scaling', 'url': reverse('media_tools:resize')},
        {'title': 'Facebook', 'icon': 'facebook', 'url': reverse('video_downloader:facebook_downloader')},
        {'title': 'X (Twitter)', 'icon': 'twitter', 'url': reverse('video_downloader:twitter_downloader')},
        {'title': 'Instagram', 'icon': 'instagram', 'url': reverse('video_downloader:instagram_downloader')},
        {'title': 'TikTok', 'icon': 'music-2', 'url': reverse('video_downloader:tiktok_downloader')},
        {'title': 'Vimeo', 'icon': 'video', 'url': reverse('video_downloader:vimeo_downloader')},
        {'title': 'Reddit', 'icon': 'hash', 'url': reverse('video_downloader:reddit_downloader')},
        {'title': 'Dailymotion', 'icon': 'play-circle', 'url': reverse('video_downloader:dailymotion_downloader')},
    ]

    link_tools = [
        {'title': 'Dynamic QR', 'icon': 'qr-code', 'url': reverse('dynamic_qr:dashboard') if is_dqr_user else reverse('dynamic_qr:login')},
        {'title': 'Short URL', 'icon': 'link', 'url': reverse('dynamic_qr:short_url') if is_dqr_user else reverse('dynamic_qr:login')},
    ]

    return {
        'pdf_tools_grouped': ordered_pdf,
        'image_tools_grouped': ordered_image,
        'video_tools': video_tools,
        'link_tools': link_tools,
        'all_tools_metadata': metadata,
    }
