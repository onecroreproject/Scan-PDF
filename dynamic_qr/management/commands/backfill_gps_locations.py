"""
Management command: backfill_gps_locations

Finds QRAnalytics records where GPS permission was granted and
coordinates exist, but country/city are still Unknown/empty.
Reverse-geocodes each record using the same reverse_geocode_coords()
function used live, then saves country/city/region/country_code.

Usage:
    python manage.py backfill_gps_locations
    python manage.py backfill_gps_locations --dry-run
    python manage.py backfill_gps_locations --delay 1.5

IMPORTANT: Uses Nominatim (OpenStreetMap). Respect their rate limit:
  - Max 1 request per second (enforced via --delay, default 1.2 s).
  - Do NOT run with --delay 0 on large datasets.
"""
import time

from django.core.management.base import BaseCommand
from django.db import transaction
from django.db.models import Q

from dynamic_qr.models import QRAnalytics
from dynamic_qr.utils import reverse_geocode_coords


class Command(BaseCommand):
    help = (
        "Backfill country/city for GPS-granted QRAnalytics records "
        "that have coordinates but no resolved location."
    )

    def add_arguments(self, parser):
        parser.add_argument(
            '--dry-run',
            action='store_true',
            default=False,
            help="Print what would be updated without saving anything.",
        )
        parser.add_argument(
            '--delay',
            type=float,
            default=1.2,
            help="Seconds to wait between geocoding requests (default: 1.2).",
        )
        parser.add_argument(
            '--limit',
            type=int,
            default=None,
            help="Maximum number of records to process (default: all).",
        )

    def handle(self, *args, **options):
        dry_run = options['dry_run']
        delay = max(options['delay'], 0.1)   # Enforce minimum 0.1 s
        limit = options['limit']

        # Find records: GPS granted + coordinates present + location not resolved
        qs = QRAnalytics.objects.filter(
            gps_permission='granted',
            gps_latitude__isnull=False,
            gps_longitude__isnull=False,
        ).filter(
            Q(country__in=['Unknown', '', 'XX']) |
            Q(country__isnull=True) |
            Q(city__in=['Unknown', '', 'Private IP']) |
            Q(city__isnull=True)
        ).order_by('id')

        total = qs.count()
        if limit:
            qs = qs[:limit]

        self.stdout.write(
            self.style.MIGRATE_HEADING(
                f"Found {total} record(s) to backfill"
                + (f" (processing first {limit})" if limit else "")
                + (" [DRY RUN]" if dry_run else "")
            )
        )

        updated = 0
        skipped = 0
        errors = 0

        for idx, record in enumerate(qs, start=1):
            lat = record.gps_latitude
            lon = record.gps_longitude

            self.stdout.write(
                f"[{idx}/{min(limit or total, total)}] "
                f"id={record.pk}  lat={lat}  lon={lon}  "
                f"current country={record.country!r}  city={record.city!r}"
            )

            try:
                geo = reverse_geocode_coords(lat, lon)
            except Exception as exc:
                self.stdout.write(self.style.ERROR(f"  → ERROR during geocoding: {exc}"))
                errors += 1
                time.sleep(delay)
                continue

            if not geo:
                self.stdout.write(self.style.WARNING("  → No result from geocoder, skipping."))
                skipped += 1
                time.sleep(delay)
                continue

            country_val = geo.get('country', '').strip()
            city_val = geo.get('city', '').strip()
            region_val = geo.get('region', '').strip()
            code_val = geo.get('country_code', '').strip().upper()

            self.stdout.write(
                f"  → Resolved: country={country_val!r}  city={city_val!r}  "
                f"region={region_val!r}  code={code_val!r}"
            )

            if not dry_run:
                update_fields = []
                if country_val:
                    record.country = country_val
                    update_fields.append('country')
                if code_val:
                    record.country_code = code_val
                    update_fields.append('country_code')
                if city_val:
                    record.city = city_val
                    update_fields.append('city')
                if region_val:
                    record.region = region_val
                    update_fields.append('region')
                # Always mark as GPS source
                record.location_source = 'gps'
                update_fields.append('location_source')

                if update_fields:
                    with transaction.atomic():
                        record.save(update_fields=update_fields)
                    updated += 1
            else:
                updated += 1  # Count as "would update"

            # Rate-limit: Nominatim allows max 1 req/second
            time.sleep(delay)

        self.stdout.write("")
        self.stdout.write(
            self.style.SUCCESS(
                f"Done. Updated={updated}  Skipped={skipped}  Errors={errors}"
                + (" [DRY RUN — no changes saved]" if dry_run else "")
            )
        )
