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
        if slug == 'images-to-video':
            continue
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
        
    # Inject images-to-video exactly as the 3rd item in 'image-tools' (Image Editing)
    if 'image-tools' in image_tools_grouped and 'images-to-video' in TOOLS:
        tool_data = {
            'title': TOOLS['images-to-video'].get('title'),
            'icon': TOOLS['images-to-video'].get('icon'),
            'slug': 'images-to-video',
            'is_coming_soon': TOOLS['images-to-video'].get('is_coming_soon', False),
            'app_name': 'converter'
        }
        # Insert at index 2 (which makes it the 3rd item)
        image_tools_grouped['image-tools']['tools'].insert(2, tool_data)

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


    # Video Tools and Link Tools for global navigation
    video_tools = [
        {'title': 'Converter', 'icon': 'video', 'url': reverse('converter:convert_page', args=['video-converter'])},
        {'title': 'Image to Video', 'icon': 'video', 'url': reverse('converter:convert_page', args=['images-to-video'])},
        {'title': 'Trim Video', 'icon': 'scissors', 'url': reverse('media_tools:trim')},
        {'title': 'Merge Video', 'icon': 'combine', 'url': reverse('media_tools:merge')},
        {'title': 'Crop Video', 'icon': 'crop', 'url': reverse('media_tools:crop')},
        {'title': 'Resize Video', 'icon': 'scaling', 'url': reverse('media_tools:resize')},
        {'title': 'Instagram', 'icon': 'instagram', 'url': reverse('media_tools:downloader_instagram')},
        {'title': 'X/Twitter', 'icon': 'twitter', 'url': reverse('media_tools:downloader_twitter')},
        {'title': 'Facebook', 'icon': 'facebook', 'url': reverse('media_tools:downloader_facebook')},
        {'title': 'Threads', 'icon': 'download', 'url': reverse('media_tools:downloader_threads')},
        {'title': 'YouTube', 'icon': 'youtube', 'url': reverse('media_tools:downloader_youtube')},
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
