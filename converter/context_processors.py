from .views import TOOLS

def tools_processor(request):
    """Make all tools available to all templates, grouped by category.

    Categories are ordered to match iLovePDF's navigation structure:
    Organize PDF → Optimize PDF → Convert to/from PDF → Edit PDF →
    PDF Security → Image Tools → Smart Creators → AI Generation →
    Utilities → Video Download
    """
    grouped_tools = {}

    # Category display names – ordered like iLovePDF
    CATEGORY_LABELS = {
        'convert': 'Convert to/from PDF',
        'pdf-tools': 'PDF Tools',
        'image-tools': 'Image Tools',
        'generate': 'Smart Creators',
        'ai-tools': 'AI Generation',
        'other': 'Utilities',
        'download': 'Video Download',
    }

    # Define the desired ordering of categories
    CATEGORY_ORDER = [
        'convert', 'pdf-tools', 'image-tools',
        'generate', 'ai-tools', 'other', 'download',
    ]

    for slug, data in TOOLS.items():
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
            'is_coming_soon': data.get('is_coming_soon', False)
        })

    # Re-order the dict based on CATEGORY_ORDER
    ordered = {}
    for cat in CATEGORY_ORDER:
        if cat in grouped_tools:
            ordered[cat] = grouped_tools[cat]
    # Append any remaining categories not in the order list
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
                'slug': slug
            }
            for slug, data in TOOLS.items()
        }
    }
