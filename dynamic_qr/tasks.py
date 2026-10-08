import logging
import json
import urllib.request
from celery import shared_task
from django.conf import settings
from .models import QRAnalytics
from .utils import is_private_address, reverse_geocode_coords

logger = logging.getLogger(__name__)

@shared_task(bind=True, max_retries=3)
def enrich_location(self, event_id, ip_address):
    """
    Asynchronously enriches a QRAnalytics event with GeoIP data.
    """
    try:
        if not ip_address or is_private_address(ip_address):
            return "Skipped (private/no IP)"
            
        event = QRAnalytics.objects.filter(pk=event_id).first()
        if not event:
            return "Event not found"

        geoip_enabled = getattr(settings, 'GEOIP_ENABLED', False)
        if not geoip_enabled:
            return "Skipped (GeoIP disabled)"

        timeout = getattr(settings, 'GEOIP_TIMEOUT', 4)
        
        # Use HTTPS if supported, else this is HTTP
        # ip-api.com pro supports HTTPS. For free tier, we still use HTTP but we log it as a privacy warning
        url = f'http://ip-api.com/json/{ip_address}?fields=status,country,countryCode,regionName,city,lat,lon'
        
        req = urllib.request.Request(url, headers={'User-Agent': 'ScanPDF-Analytics/1.0'})
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            geo_data = json.loads(resp.read().decode())
            if geo_data.get('status') == 'success':
                updates = {
                    'country': geo_data.get('country', 'Unknown'),
                    'country_code': geo_data.get('countryCode', 'XX'),
                    'region': geo_data.get('regionName', 'Unknown'),
                    'city': geo_data.get('city', 'Unknown'),
                    'latitude': geo_data.get('lat'),
                    'longitude': geo_data.get('lon'),
                    'location_source': 'ip'
                }
                
                # J9: Data Validation before saving
                # Validate bounds and types
                lat = updates['latitude']
                lon = updates['longitude']
                if lat is not None and not (-90 <= lat <= 90):
                    updates['latitude'] = None
                if lon is not None and not (-180 <= lon <= 180):
                    updates['longitude'] = None
                
                # Update only if event hasn't been enriched by GPS already
                if event.location_source not in ('gps', 'local'):
                    QRAnalytics.objects.filter(pk=event_id).update(**updates)
                    return f"Enriched event {event_id}"
                return "Skipped (already enriched by GPS/local)"
            return "GeoIP provider returned failure"
            
    except Exception as exc:
        logger.warning(f"Failed to enrich location for event {event_id}: {exc}")
        self.retry(exc=exc, countdown=10)


@shared_task(bind=True, max_retries=2)
def reverse_geocode_gps(self, event_id, latitude, longitude):
    """
    Asynchronously enriches a GPS-authorized QRAnalytics event with human-readable location data.
    """
    try:
        # Validate coordinates
        if not (-90 <= latitude <= 90 and -180 <= longitude <= 180):
            return "Invalid coordinates"

        event = QRAnalytics.objects.filter(pk=event_id).first()
        if not event:
            return "Event not found"

        geo = reverse_geocode_coords(latitude, longitude)
        if geo:
            updates = {}
            if geo.get('country', '').strip():
                updates['country'] = geo['country'].strip()
            if geo.get('city', '').strip():
                updates['city'] = geo['city'].strip()
            if geo.get('region', '').strip():
                updates['region'] = geo['region'].strip()
            if geo.get('country_code', '').strip():
                updates['country_code'] = geo['country_code'].strip().upper()
                
            if updates:
                QRAnalytics.objects.filter(pk=event_id).update(**updates)
                return f"Reverse geocoded event {event_id}"
        return "No geocoding data found"
        
    except Exception as exc:
        logger.warning(f"Failed to reverse geocode event {event_id}: {exc}")
        self.retry(exc=exc, countdown=15)
