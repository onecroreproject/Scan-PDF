import os
from django.conf import settings
from django.template.loader import render_to_string
from django.http import HttpResponse

class CustomErrorPageMiddleware:
    def __init__(self, get_response):
        self.get_response = get_response
        self.supported_errors = [400, 401, 403, 404, 405, 408, 409, 410, 413, 415, 422, 429, 500, 501]

    def __call__(self, request):
        response = self.get_response(request)

        # Do not interfere with API/JSON responses
        if response.get('Content-Type', '').startswith('application/json'):
            return response

        # Do not override explicitly rendered HTML unless it's the default HttpResponseNotAllowed etc.
        # Actually, Django's default HttpResponseNotAllowed has no content or minimal content.
        # It's safest to check if the status is in our list and the content is empty or short (default django error text)
        status = response.status_code
        
        # Django automatically handles 404 and 500 with our templates, so skip them to avoid double rendering
        if status in [404, 500]:
            return response

        if status in self.supported_errors:
            # If it's a standard Django empty/minimal response for these codes, replace it with our template
            content = response.content.decode('utf-8', errors='ignore').lower()
            if len(content) < 500 or "method not allowed" in content or "forbidden" in content or "bad request" in content:
                template_name = f"{status}.html"
                try:
                    rendered = render_to_string(template_name, request=request)
                    return HttpResponse(rendered, status=status)
                except Exception:
                    pass

        return response
