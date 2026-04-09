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
        'download': 'Video Download',
        'audio-tools': 'Audio Editor',
    }

    CATEGORY_ORDER = [
        'convert', 'pdf-tools', 'image-tools', 'image-pro', 'image-conv',
        'generate', 'ai-tools', 'other', 'download', 'audio-tools'
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

        grouped_tools[cat]['tools'].append({
            'title': data.get('title'),
            'icon': data.get('icon'),
            'slug': slug,
            'is_coming_soon': data.get('is_coming_soon', False),
            'app_name': 'image_processor' if slug in IMAGE_TOOLS else 'converter'
        })

    # Re-order the dict
    ordered = {}
    for cat in CATEGORY_ORDER:
        if cat in grouped_tools:
            ordered[cat] = grouped_tools[cat]
    for cat, info in grouped_tools.items():
        if cat not in ordered:
            ordered[cat] = info

    return {
        'grouped_tools': ordered,
        'all_tools_metadata': {
            slug: {
                'title': data['title'],
                'icon': data['icon'],
                'description': data.get('description', ''),
                'slug': slug,
                'url': reverse('image_processor:tool_page', args=[slug]) if slug in IMAGE_TOOLS else reverse('converter:convert_page', args=[slug])
            }
            for slug, data in all_combined.items()
        }
    }
