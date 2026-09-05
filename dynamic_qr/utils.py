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

SOURCE_LABELS = {
    'direct': 'Direct Visit',
    'internal': 'Internal Navigation',
    'qr': 'QR Scan',
    'search': 'Search Engine',
    'social': 'Social Media',
    'referral': 'Referral Website',
    'unknown': 'Unknown',
}


def classify_traffic_source(referrer, is_qr_scan):
    if is_qr_scan:
        return 'qr'
    if not referrer:
        return 'direct'
    ref_lower = referrer.lower()
    if any(value in ref_lower for value in ('google', 'bing', 'yahoo', 'duckduckgo', 'baidu', 'yandex')):
        return 'search'
    if any(value in ref_lower for value in (
        'facebook', 'instagram', 'twitter', 't.co', 'linkedin', 'x.com',
        'tiktok', 'snapchat', 'whatsapp', 'telegram', 'reddit', 'pinterest',
    )):
        return 'social'
    if any(value in ref_lower for value in ('scanpdf', '127.0.0.1', 'localhost')):
        return 'internal'
    return 'referral'


def is_private_address(value):
    try:
        return ipaddress.ip_address(value).is_private or ipaddress.ip_address(value).is_loopback
    except ValueError:
        return False

def record_short_url_event(qr, request, result, status, visitor_id=None):
    """
    Centralized analytics recording pipeline.
    Must be called for every short URL hit, regardless of outcome.
    """
    # 1. Parse User Agent & Bot detection
    ua = request.META.get('HTTP_USER_AGENT', '').lower()
    bot_keywords = ['bot', 'crawl', 'spider', 'slurp', 'mediapartners', 'preview', 'slack', 'discord', 'whatsapp', 'skype']
    is_bot = any(b in ua for b in bot_keywords)
    if is_bot and result == 'redirect_success':
        # Re-classify successful requests from bots as bot_request
        result = 'bot_request'
        
    # 2. Extract Real IP
    x_forwarded_for = request.META.get('HTTP_X_FORWARDED_FOR')
    if x_forwarded_for:
        ip = x_forwarded_for.split(',')[0].strip()
    else:
        ip = request.META.get('REMOTE_ADDR', '')

    # 3. Handle Traffic Source & QR tracking
    is_qr_scan = request.GET.get('source') == 'qr'
    referrer = request.META.get('HTTP_REFERER', '')[:500]
    
    # Classify source
    source = classify_traffic_source(referrer, is_qr_scan)

    # 4. Generate stable visitor ID if not provided
    if not visitor_id:
        ip_base = ip.rsplit('.', 1)[0] if '.' in ip else ip
        visitor_string = f"{ip_base}_{ua}_{qr.id}"
        visitor_id = hashlib.sha256(visitor_string.encode('utf-8')).hexdigest()[:32]

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
    
    # We only process Geo IP if it's not a bot to save external API limits
    if ip and not is_bot:
        is_private = is_private_address(ip)
        if not is_private:
            try:
                headers = {'User-Agent': 'ScanPDF/1.0'}
                req = Request(f'http://ip-api.com/json/{ip}?fields=status,country,countryCode,regionName,city,lat,lon', headers=headers)
                with urlopen(req, timeout=4) as resp:
                    geo_data = json.loads(resp.read().decode())
                    if geo_data.get('status') == 'success':
                        country = geo_data.get('country', 'Unknown')
                        country_code = geo_data.get('countryCode', 'XX')
                        region = geo_data.get('regionName', 'Unknown')
                        city = geo_data.get('city', 'Unknown')
                        lat = geo_data.get('lat')
                        lon = geo_data.get('lon')
            except Exception:
                pass
        else:
            country, country_code, region, city = 'Unknown', 'XX', 'Unknown', 'Unknown'

    location_source = 'local' if is_private_address(ip) else ('ip' if country != 'Unknown' else 'unknown')

    # If this request came through the GPS allow flow, override lat/lon
    # We pass gps_lat and gps_lon in request.session if it's authorized
    gps_lat = request.session.pop(f'qr_gps_lat_{qr.id}', None)
    gps_lon = request.session.pop(f'qr_gps_lon_{qr.id}', None)
    if gps_lat and gps_lon:
        lat = float(gps_lat)
        lon = float(gps_lon)

    # 7. Record analytics and the cached successful-click counter together.
    try:
        with transaction.atomic():
            QRAnalytics.objects.create(
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
                location_source=location_source,
                gps_permission='not_required',
                redirect_result=result,
                http_status=status
            )
            if result == 'redirect_success':
                type(qr).objects.filter(pk=qr.pk).update(scan_count=F('scan_count') + 1)
    except Exception:
        logger.exception("Unable to record short URL event for %s", qr.pk)


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
        updates.update(redirect_result='redirect_success', http_status=302, location_source='gps')
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
    request.session.pop(f'qr_pending_event_{qr.id}', None)
    request.session[f'qr_gps_auth_{qr.id}'] = permission == 'granted'
    return event

