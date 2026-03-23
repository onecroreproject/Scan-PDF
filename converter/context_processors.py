from .views import TOOLS

def tools_processor(request):
    """Make all tools available to all templates, grouped by category."""
    grouped_tools = {}
    
    # Category Display Names
    CATEGORY_LABELS = {
        'convert': 'Convert to/from PDF',
        'pdf-tools': 'Optimize & Org',
        'image-tools': 'Image Tools',
        'generate': 'Smart Creators',
        'ai-tools': 'AI Generation',
        'other': 'Utilities',
        'download': 'Video Download'
    }

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
        
    return {
        'grouped_tools': grouped_tools,
        'all_tools_metadata': {slug: {'title': data['title'], 'icon': data['icon'], 'description': data.get('description', ''), 'slug': slug} for slug, data in TOOLS.items()}
    }
