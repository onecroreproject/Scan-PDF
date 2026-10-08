import re
import socket
import ipaddress
from urllib.parse import urlparse
from django.core.exceptions import ValidationError

ALLOWED_PROTOCOLS = {'http', 'https', 'mailto', 'tel', 'sms', 'geo', 'wifi'}
BLOCKED_PROTOCOLS = {'javascript', 'data', 'file', 'vbscript'}

def validate_protocol(url):
    """Ensure the URL uses an allowed protocol and explicitly blocks dangerous ones."""
    if not url:
        return
    parsed = urlparse(url)
    scheme = (parsed.scheme or '').lower()
    if scheme in BLOCKED_PROTOCOLS:
        raise ValidationError(f"Dangerous protocol '{scheme}' is strictly forbidden.")
    # For redirect targets, we accept the ALLOWED_PROTOCOLS
    # (If scheme is missing, it's relative, but we assume http/https later).
    if scheme and scheme not in ALLOWED_PROTOCOLS:
        raise ValidationError(f"Protocol '{scheme}' is not supported.")

def validate_redirect_target(url):
    """
    Validates a URL meant for CLIENT-SIDE redirects.
    This allows private IPs (for testing/intranets) but blocks dangerous schemes.
    """
    validate_protocol(url)

def get_ip_from_hostname(hostname):
    try:
        return socket.gethostbyname(hostname)
    except socket.gaierror:
        return None

def validate_ssrf_safe(url):
    """
    Validates a URL meant for SERVER-SIDE fetching.
    Strictly forbids private IPs, loopback, broadcast, etc.
    Returns the resolved IP address to be used for the connection to prevent DNS rebinding.
    """
    if not url:
        return None, url
    parsed = urlparse(url)
    if parsed.scheme not in ('http', 'https'):
        raise ValidationError("Only HTTP/HTTPS allowed for server-side requests.")

    hostname = parsed.hostname
    if not hostname:
        raise ValidationError("Invalid URL format.")

    # Always resolve the hostname to an IP to prevent TOCTOU / DNS rebinding
    resolved_ip = get_ip_from_hostname(hostname)
    if not resolved_ip:
        raise ValidationError(f"Could not resolve hostname: {hostname}")

    try:
        ip_obj = ipaddress.ip_address(resolved_ip)
        if ip_obj.is_private or ip_obj.is_loopback or ip_obj.is_link_local or ip_obj.is_reserved or ip_obj.is_multicast:
            raise ValidationError("Target resolves to a restricted private or local IP.")
    except ValueError:
        raise ValidationError("Invalid resolved IP.")

    # Rebuild URL with the exact IP so the subsequent request connects to this validated IP
    # (Note: for SNI/Host headers to work, we must supply the original hostname in headers)
    safe_url = parsed._replace(netloc=f"{resolved_ip}:{parsed.port}" if parsed.port else resolved_ip).geturl()
    return resolved_ip, safe_url

def validate_header(header_value):
    """Validate header format and length."""
    if not header_value:
        return
    if not re.match(r'^[A-Za-z0-9_-]{1,20}$', header_value):
        raise ValidationError("Header must be 1-20 characters (A-Z, 0-9, -, _).")
