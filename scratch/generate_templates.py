import os

templates = {
    "400": ("400", "alert-circle", "Bad Request", "The server cannot process the request due to a client error."),
    "401": ("401", "lock", "Unauthorized", "You must authenticate to access this page."),
    "403": ("403", "shield-alert", "Access Denied", "You don't have permission to view this resource."),
    "404": ("404", "map-pin-off", "Page Not Found", "The page you're looking for doesn't exist or has been moved."),
    "405": ("405", "slash", "Method Not Allowed", "The HTTP method used is not allowed for this resource."),
    "408": ("408", "clock", "Request Timeout", "The server timed out waiting for the request."),
    "409": ("409", "git-merge", "Conflict", "The request could not be completed due to a conflict with the current state of the resource."),
    "410": ("410", "trash-2", "Gone", "This resource has been permanently removed and is no longer available."),
    "413": ("413", "file-warning", "Payload Too Large", "The uploaded file exceeds our maximum allowed size."),
    "415": ("415", "file-question", "Unsupported Media Type", "The requested format is not supported by the server."),
    "422": ("422", "file-x", "Unprocessable Entity", "The request was formatted correctly but contains invalid data."),
    "429": ("429", "gauge", "Too Many Requests", "You've sent too many requests in a given amount of time. Please slow down."),
    "500": ("500", "server-crash", "Internal Server Error", "Our servers encountered an unexpected error. We are looking into it."),
    "501": ("501", "construction", "Not Implemented", "This feature is currently not available or not implemented."),
}

template_content = """{{% extends 'errors/base.html' %}}

{{% block error_title %}}{title}{{% endblock %}}
{{% block error_code %}}{code}{{% endblock %}}
{{% block error_icon %}}<i data-lucide="{icon}" class="w-12 h-12 text-white"></i>{{% endblock %}}
{{% block error_heading %}}{title}{{% endblock %}}
{{% block error_description %}}{desc}{{% endblock %}}
"""

os.makedirs(r"r:\DLK-Scan-PDF\Scan-PDF\templates", exist_ok=True)
for code, (c, icon, title, desc) in templates.items():
    with open(rf"r:\DLK-Scan-PDF\Scan-PDF\templates\{code}.html", "w") as f:
        f.write(template_content.format(title=title, code=c, icon=icon, desc=desc))

print("Created Django error templates.")
