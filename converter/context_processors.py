from .views import TOOLS

def tools_processor(request):
    """Make all tools available to all templates with only necessary fields."""
    sanitized_tools = {}
    for slug, data in TOOLS.items():
        sanitized_tools[slug] = {
            'title': data.get('title'),
            'description': data.get('description'),
            'icon': data.get('icon'),
            'slug': slug
        }
    return {'all_tools_metadata': sanitized_tools}
