import hashlib
import ipaddress
from urllib.request import urlopen, Request
import json
import logging
from django.db import transaction
from django.db.models import F
from django.utils import timezone
from .models import QRAnalytics

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Reverse geocoding — GPS coordinates → country / city / region
# ---------------------------------------------------------------------------

def _nominatim_reverse_http(latitude, longitude, timeout=5):
    """
    Pure-HTTP Nominatim reverse geocode (no geopy dependency required).
    Returns a dict with any subset of: country, country_code, city, region.
    Returns None on any failure.
    """
    url = (
        f'https://nominatim.openstreetmap.org/reverse'
        f'?lat={latitude}&lon={longitude}&format=json&addressdetails=1&accept-language=en'
    )
    headers = {
        'User-Agent': 'ScanPDF-ShortURL-Analytics/1.0 (contact@scanpdf.io)',
        'Accept-Language': 'en',
    }
    try:
        req = Request(url, headers=headers)
        with urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode('utf-8'))
        address = data.get('address', {})
        if not address:
            return None

        country = (address.get('country') or '').strip()
        country_code = (address.get('country_code') or '').strip().upper()
        region = (address.get('state') or address.get('region') or '').strip()
        city = (
            address.get('city')
            or address.get('town')
            or address.get('village')
            or address.get('municipality')
            or address.get('suburb')
            or address.get('county')
            or address.get('state_district')
            or ''
        ).strip()

        result = {}
        if country:
            result['country'] = country
        if country_code:
            result['country_code'] = country_code
        if region:
            result['region'] = region
        if city:
            result['city'] = city
        return result if result else None
    except Exception:
        return None


def reverse_geocode_coords(latitude, longitude):
    """
    Reverse-geocode GPS coordinates to country/city/region.

    Strategy:
      1. Try geopy Nominatim (if installed).
      2. Fall back to direct HTTP call to Nominatim via urllib.

    Returns a dict with any subset of:
        { 'country', 'country_code', 'city', 'region' }
    Returns None on failure — caller must handle gracefully.
    """
    if not (-90 <= latitude <= 90 and -180 <= longitude <= 180):
        return None

    # Attempt 1: geopy Nominatim (preferred)
    try:
        from geopy.geocoders import Nominatim

        geolocator = Nominatim(
            user_agent='ScanPDF-ShortURL-Analytics/1.0',
        )
        location = geolocator.reverse(
            (latitude, longitude),
            language='en',
            exactly_one=True,
            timeout=5,
        )
        if location and location.raw:
            address = location.raw.get('address', {})
            country = (address.get('country') or '').strip()
            country_code = (address.get('country_code') or '').strip().upper()
            region = (address.get('state') or address.get('region') or '').strip()
            city = (
                address.get('city')
                or address.get('town')
                or address.get('village')
                or address.get('municipality')
                or address.get('suburb')
                or address.get('county')
                or address.get('state_district')
                or ''
            ).strip()

            result = {}
            if country:
                result['country'] = country
            if country_code:
                result['country_code'] = country_code
            if region:
                result['region'] = region
            if city:
                result['city'] = city
            return result if result else None

    except ImportError:
        # geopy not installed — fall through to HTTP fallback
        pass
    except Exception:
        # Timeout / service error — fall through to HTTP fallback
        pass

    # Attempt 2: Pure urllib HTTP fallback
    return _nominatim_reverse_http(latitude, longitude, timeout=5)


# ---------------------------------------------------------------------------

SOURCE_LABELS = {
    'direct': 'Direct Visit',
    'internal': 'Internal Navigation',
    'qr': 'QR Scan',
    'search': 'Search Engine',
    'social': 'Social Media',
    'referral': 'Referral Website',
    'unknown': 'Unknown',
}


def detect_source(referrer, is_qr_scan):
    """
    Returns (detected_source, detected_medium).
    """
    if is_qr_scan:
        return ('QR Scan', 'qr')
    if not referrer:
        return ('Direct / Unknown', 'direct')
    
    ref_lower = referrer.lower()
    
    # Specific Platforms
    if 'facebook.com' in ref_lower or 'l.facebook.com' in ref_lower or 'm.facebook.com' in ref_lower:
        return ('Facebook', 'social')
    if 'instagram.com' in ref_lower:
        return ('Instagram', 'social')
    if 'youtube.com' in ref_lower or 'youtu.be' in ref_lower:
        return ('YouTube', 'social')
    if 'linkedin.com' in ref_lower:
        return ('LinkedIn', 'social')
    if 'twitter.com' in ref_lower or 't.co' in ref_lower or 'x.com' in ref_lower:
        return ('X (Twitter)', 'social')
    if 'tiktok.com' in ref_lower:
        return ('TikTok', 'social')
    if 'pinterest.com' in ref_lower:
        return ('Pinterest', 'social')
    if 'snapchat.com' in ref_lower:
        return ('Snapchat', 'social')
    if 'reddit.com' in ref_lower:
        return ('Reddit', 'social')
    if 't.me' in ref_lower or 'telegram.org' in ref_lower:
        return ('Telegram', 'social')
    if 'whatsapp.com' in ref_lower or 'wa.me' in ref_lower:
        return ('WhatsApp', 'social')
        
    # Generic Search
    if any(s in ref_lower for s in ('google.com', 'google.co')):
        return ('Google', 'search')
    if 'bing.com' in ref_lower:
        return ('Bing', 'search')
    if 'yahoo.com' in ref_lower:
        return ('Yahoo', 'search')
    if 'duckduckgo.com' in ref_lower:
        return ('DuckDuckGo', 'search')
    if 'baidu.com' in ref_lower:
        return ('Baidu', 'search')
    if 'yandex.ru' in ref_lower or 'yandex.com' in ref_lower:
        return ('Yandex', 'search')

    if any(value in ref_lower for value in ('scanpdf', '127.0.0.1', 'localhost')):
        return ('Internal Navigation', 'internal')
        
    from urllib.parse import urlparse
    parsed = urlparse(referrer)
    domain = parsed.netloc if parsed.netloc else referrer[:40]
    return (domain, 'referral')


def is_private_address(value):
    try:
        return ipaddress.ip_address(value).is_private or ipaddress.ip_address(value).is_loopback
    except ValueError:
        return False

def get_client_ip(request):
    """
    Securely determines the client's real IP address.
    Only processes X-Forwarded-For if the immediate upstream (REMOTE_ADDR) is a trusted proxy.
    Walks the chain backwards and discards trusted proxies until it hits the first untrusted IP.
    """
    from django.conf import settings
    import ipaddress

    # Allow configuration via settings, default to empty (trust no one)
    trusted_proxies_cfg = getattr(settings, 'TRUSTED_PROXY_IPS', [])
    trusted_networks = []
    
    for proxy in trusted_proxies_cfg:
        try:
            # Handle both single IPs and CIDR networks
            if '/' in proxy:
                trusted_networks.append(ipaddress.ip_network(proxy, strict=False))
            else:
                trusted_networks.append(ipaddress.ip_network(f"{proxy}/32" if ':' not in proxy else f"{proxy}/128", strict=False))
        except ValueError:
            pass # Ignore malformed config safely

    def is_trusted(ip_str):
        try:
            ip_obj = ipaddress.ip_address(ip_str.strip())
            return any(ip_obj in net for net in trusted_networks)
        except ValueError:
            return False

    remote_addr = request.META.get('REMOTE_ADDR', '').strip()
    if not remote_addr:
        return '0.0.0.0' # Fallback safely if impossible
        
    x_forwarded_for = request.META.get('HTTP_X_FORWARDED_FOR', '')

    # Default to safe behavior
    if not x_forwarded_for or not is_trusted(remote_addr):
        return remote_addr

    # Parse X-Forwarded-For: client, proxy1, proxy2
    chain = [ip.strip() for ip in x_forwarded_for.split(',')]
    
    # We read from right (closest proxy) to left (client)
    # The immediate upstream is remote_addr (already verified as trusted)
    # The last element in the chain is proxy2.
    # We walk backwards. The first IP that is NOT a trusted proxy is the real client.
    for ip_str in reversed(chain):
        if not is_trusted(ip_str):
            # Must validate it's a real IP address before returning it
            try:
                ipaddress.ip_address(ip_str)
                return ip_str
            except ValueError:
                break # Malformed IP in chain, fall back to last known safe

    return remote_addr


def get_visitor_id(request, qr, client_ip=None):
    """Generates a stable, privacy-preserving visitor identifier using HMAC."""
    from django.conf import settings
    import hmac
    
    if not client_ip:
        client_ip = get_client_ip(request)
    ua = request.META.get('HTTP_USER_AGENT', '').lower()
    ip_base = client_ip.rsplit('.', 1)[0] if '.' in client_ip else client_ip
    visitor_string = f"{ip_base}_{ua}_{qr.id}"
    
    secret = getattr(settings, 'SECRET_KEY', 'fallback-secret').encode('utf-8')
    return hmac.new(secret, visitor_string.encode('utf-8'), hashlib.sha256).hexdigest()[:32]


def is_duplicate_request(visitor_id, window_seconds=5):
    """
    Redis-backed atomic dedupe for Short URL analytics success scans.
    Returns True if the request is a duplicate within the window.
    """
    from django.core.cache import cache
    # I-CHECK 4, 5: Removed raw IP, visitor_id already hashes IP/UA/QR.
    raw_key = f"shorturl:success:{visitor_id}"
    key_hash = hashlib.sha256(raw_key.encode('utf-8')).hexdigest()
    cache_key = f"shorturl:success:{key_hash}"
    
    try:
        return not cache.add(cache_key, 1, timeout=window_seconds)
    except Exception as e:
        logger.warning(f"Analytics dedupe cache failed, failing open: {e}")
        return False


def should_record_failure_event(visitor_id, event_type, window_seconds=60):
    """
    Redis-backed atomic dedupe for Short URL failure events (e.g. disabled, expired, password_failed).
    Returns True if we SHOULD record the event (i.e. it is NOT a duplicate within the window).
    """
    from django.core.cache import cache
    raw_key = f"shorturl:failure:{event_type}:{visitor_id}"
    key_hash = hashlib.sha256(raw_key.encode('utf-8')).hexdigest()
    cache_key = f"shorturl:failure:{key_hash}"
    
    try:
        return cache.add(cache_key, 1, timeout=window_seconds)
    except Exception as e:
        logger.warning(f"Failure analytics dedupe cache failed, failing open: {e}")
        return True


def record_short_url_event(qr, request, result, status, visitor_id=None, utm_data=None, was_cloaked=False):
    """
    Centralized analytics recording pipeline.
    Must be called for every short URL hit, regardless of outcome.
    """
    incoming_utm = {
        'utm_source': (utm_data or {}).get('utm_source') if utm_data else None,
        'utm_medium': (utm_data or {}).get('utm_medium') if utm_data else None,
        'utm_campaign': (utm_data or {}).get('utm_campaign') if utm_data else None,
        'utm_term': (utm_data or {}).get('utm_term') if utm_data else None,
        'utm_content': (utm_data or {}).get('utm_content') if utm_data else None,
    }
    configured_utm = {
        'utm_source': qr.utm_source,
        'utm_medium': qr.utm_medium,
        'utm_campaign': qr.utm_campaign,
        'utm_term': qr.utm_term,
        'utm_content': qr.utm_content,
    }
    # 1. Parse User Agent & Bot detection
    ua = request.META.get('HTTP_USER_AGENT', '').lower()
    bot_keywords = ['bot', 'crawl', 'spider', 'slurp', 'mediapartners', 'preview', 'slack', 'discord', 'whatsapp', 'skype']
    is_bot = any(b in ua for b in bot_keywords)
    if is_bot and result == 'redirect_success':
        # Re-classify successful requests from bots as bot_request
        result = 'bot_request'
        
    # 2. Extract Real IP using secure helper
    ip = get_client_ip(request)

    # 3. Handle Traffic Source & QR tracking
    is_qr_scan = request.GET.get('source') == 'qr'
    referrer = request.META.get('HTTP_REFERER', '')[:500]
    
    # Classify source
    detected_source, detected_medium = detect_source(referrer, is_qr_scan)
    # Map back to old source field for backward compatibility
    source = detected_medium

    # 4. Generate stable visitor ID if not provided
    if not visitor_id:
        visitor_id = get_visitor_id(request, qr, ip)

    # 5. Extract Tech Specs
    browser = 'Other'
    if 'edg/' in ua or 'edge' in ua: browser = 'Edge'
    elif 'samsungbrowser' in ua: browser = 'Samsung Internet'
    elif 'opera' in ua or 'opr/' in ua: browser = 'Opera'
    elif 'chrome' in ua and 'safari' in ua: browser = 'Chrome'
    elif 'safari' in ua and 'chrome' not in ua: browser = 'Safari'
    elif 'firefox' in ua: browser = 'Firefox'
    
    os_name = 'Other'
    if 'windows' in ua: os_name = 'Windows'
    elif 'iphone' in ua or 'ipad' in ua: os_name = 'iOS'
    elif 'mac' in ua: os_name = 'macOS'
    elif 'android' in ua: os_name = 'Android'
    elif 'linux' in ua: os_name = 'Linux'
    
    device = 'Desktop'
    if 'ipad' in ua or 'tablet' in ua or ('android' in ua and 'mobile' not in ua):
        device = 'Tablet'
    elif 'mobile' in ua or 'iphone' in ua or 'android' in ua:
        device = 'Mobile'
    
    # 6. Extract Geolocations
    country, country_code, region, city = 'Unknown', 'XX', 'Unknown', 'Unknown'
    lat, lon = None, None
    location_source = 'local' if is_private_address(ip) else 'unknown'

    # If this request came through the GPS allow flow, override lat/lon
    # We pass gps_lat and gps_lon in request.session if it's authorized
    gps_lat = request.session.pop(f'qr_gps_lat_{qr.id}', None)
    gps_lon = request.session.pop(f'qr_gps_lon_{qr.id}', None)
    if gps_lat and gps_lon:
        lat = float(gps_lat)
        lon = float(gps_lon)
        location_source = 'gps'

    # 7. Record analytics and the cached successful-click counter together.
    try:
        with transaction.atomic():
            event = QRAnalytics.objects.create(
                qr_code=qr,
                ip_address=ip,
                user_agent=ua[:500],
                browser=browser,
                os=os_name,
                device_type=device,
                country=country,
                country_code=country_code,
                region=region,
                city=city,
                latitude=lat,
                longitude=lon,
                referrer=referrer,
                is_bot=is_bot,
                is_qr_scan=is_qr_scan,
                source=source,
                visitor_id=visitor_id,
                detected_source=detected_source,
                detected_medium=detected_medium,
                location_source=location_source,
                gps_permission='not_required',
                redirect_result=result,
                http_status=status,
                utm_source=configured_utm['utm_source'],
                utm_medium=configured_utm['utm_medium'],
                utm_campaign=configured_utm['utm_campaign'],
                utm_term=configured_utm['utm_term'],
                utm_content=configured_utm['utm_content'],
                incoming_utm_source=incoming_utm['utm_source'],
                incoming_utm_medium=incoming_utm['utm_medium'],
                incoming_utm_campaign=incoming_utm['utm_campaign'],
                incoming_utm_term=incoming_utm['utm_term'],
                incoming_utm_content=incoming_utm['utm_content'],
                was_cloaked=was_cloaked
            )
            if result == 'redirect_success':
                type(qr).objects.filter(pk=qr.pk).update(scan_count=F('scan_count') + 1)
                
            # J23, J24: Enqueue async GeoIP lookup after commit
            if ip and not is_bot and location_source == 'unknown':
                transaction.on_commit(lambda e_id=event.id, cip=ip: _enqueue_geoip(e_id, cip))
    except Exception:
        logger.exception("Unable to record short URL event for %s", qr.pk)

def _enqueue_geoip(event_id, ip_address):
    try:
        from .tasks import enrich_location
        enrich_location.delay(event_id, ip_address)
    except Exception as e:
        logger.warning(f"Failed to enqueue GeoIP task for event {event_id}: {e}")


def update_pending_gps_event(request, qr, permission, latitude=None, longitude=None, accuracy=None):
    """Update the session-bound GPS event; never trust a client-supplied event id."""
    event_id = request.session.get(f'qr_pending_event_{qr.id}')
    if not event_id:
        return None
    if permission == 'granted':
        if not (-90 <= latitude <= 90 and -180 <= longitude <= 180):
            raise ValueError('GPS coordinates are outside valid ranges.')
        if accuracy is None or accuracy < 0:
            raise ValueError('GPS accuracy must be zero or greater.')
    updates = {
        'gps_permission': permission,
        'gps_latitude': latitude,
        'gps_longitude': longitude,
        'gps_accuracy': accuracy,
        'gps_captured_at': timezone.now() if permission == 'granted' else None,
    }
    if permission == 'granted':
        updates.update(
            redirect_result='redirect_success',
            http_status=302,
            location_source='gps',
            latitude=latitude,
            longitude=longitude,
        )

        # J16, J18: Reverse-geocode GPS coordinates to country/city ASYNCHRONOUSLY
        # Enqueue the Celery task after the atomic commit block.

    elif permission in ('denied', 'unavailable', 'timeout'):
        updates.update(redirect_result='gps_denied', http_status=403)

    with transaction.atomic():
        event = QRAnalytics.objects.select_for_update().filter(
            pk=event_id, qr_code=qr, redirect_result='gps_required'
        ).first()
        if not event:
            return None
        for field, value in updates.items():
            setattr(event, field, value)
        event.save(update_fields=list(updates))
        if permission == 'granted':
            type(qr).objects.filter(pk=qr.pk).update(scan_count=F('scan_count') + 1)
            transaction.on_commit(lambda e_id=event.id, lat=latitude, lon=longitude: _enqueue_reverse_geocode(e_id, lat, lon))
    request.session.pop(f'qr_pending_event_{qr.id}', None)
    request.session[f'qr_gps_auth_{qr.id}'] = permission == 'granted'
    return event

def _enqueue_reverse_geocode(event_id, latitude, longitude):
    try:
        from .tasks import reverse_geocode_gps
        reverse_geocode_gps.delay(event_id, latitude, longitude)
    except Exception as e:
        logger.warning(f"Failed to enqueue reverse geocode task for event {event_id}: {e}")

