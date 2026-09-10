from unittest.mock import patch
import json

from django.contrib.auth import get_user_model
from django.http import HttpResponse
from django.test import Client, TestCase
from django.utils import timezone

from .models import DynamicQRCode, QRAnalytics
from . import utils


class ShortURLAnalyticsTests(TestCase):
    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username='analytics-owner', password='password'
        )
        self.qr = DynamicQRCode.objects.create(
            user=self.user,
            qr_name='Tracked link',
            qr_type='custom-url',
            destination_url='https://example.com/destination',
            qr_enabled=True,
        )
        self.client = Client()
        self.user_agent = (
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
            'AppleWebKit/537.36 Chrome/125.0.0.0 Safari/537.36'
        )

    def test_successful_redirect_records_click_and_metadata(self):
        response = self.client.get(
            f'/qr/r/{self.qr.short_code}/',
            HTTP_USER_AGENT=self.user_agent,
        )

        self.assertEqual(response.status_code, 302)
        self.qr.refresh_from_db()
        event = QRAnalytics.objects.get(qr_code=self.qr)
        self.assertEqual(self.qr.scan_count, 1)
        self.assertEqual(event.redirect_result, 'redirect_success')
        self.assertEqual(event.browser, 'Chrome')
        self.assertEqual(event.os, 'Windows')
        self.assertEqual(event.device_type, 'Desktop')
        self.assertEqual(event.source, 'direct')
        self.assertEqual(event.location_source, 'local')
        self.assertEqual(event.country, 'Unknown')
        self.assertEqual(event.city, 'Unknown')

    def test_qr_source_and_unique_visitor_count(self):
        for source in ('qr', 'qr'):
            Client().get(
                f'/qr/r/{self.qr.short_code}/?source={source}',
                HTTP_USER_AGENT=self.user_agent,
                REMOTE_ADDR='127.0.0.1',
            )

        events = QRAnalytics.objects.filter(qr_code=self.qr)
        self.assertEqual(events.filter(is_qr_scan=True).count(), 2)
        self.assertEqual(events.filter(source='qr').count(), 2)
        self.assertEqual(events.values('visitor_id').distinct().count(), 1)

    def test_gps_permission_updates_one_event(self):
        self.qr.require_gps = True
        self.qr.save(update_fields=['require_gps'])

        with patch('dynamic_qr.views.render', return_value=HttpResponse('')):
            prompt = self.client.get(f'/qr/r/{self.qr.short_code}/')
        self.assertEqual(prompt.status_code, 200)
        pending = QRAnalytics.objects.get(qr_code=self.qr)
        self.assertEqual(pending.gps_permission, 'pending')

        response = self.client.post(
            f'/qr/r/{self.qr.short_code}/',
            {'gps_lat': '13.0827', 'gps_lon': '80.2707', 'gps_accuracy': '12.5'},
        )
        self.assertEqual(response.status_code, 302)
        follow_up = self.client.get(response['Location'])
        self.assertEqual(follow_up.status_code, 302)
        self.qr.refresh_from_db()
        pending.refresh_from_db()
        self.assertEqual(QRAnalytics.objects.filter(qr_code=self.qr).count(), 1)
        self.assertEqual(self.qr.scan_count, 1)
        self.assertEqual(pending.gps_permission, 'granted')
        self.assertEqual(pending.location_source, 'gps')
        self.assertEqual(pending.gps_accuracy, 12.5)

    def test_gps_json_post_returns_destination_json(self):
        self.qr.require_gps = True
        self.qr.save(update_fields=['require_gps'])
        with patch('dynamic_qr.views.render', return_value=HttpResponse('')):
            self.client.get(f'/qr/r/{self.qr.short_code}/')

        response = self.client.post(
            f'/qr/r/{self.qr.short_code}/',
            data=json.dumps({
                'latitude': 13.0827,
                'longitude': 80.2707,
                'accuracy': 12.5,
                'permission': 'granted',
            }),
            content_type='application/json',
        )

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json(), {
            'success': True,
            'redirect_url': 'https://example.com/destination',
        })
        self.assertEqual(QRAnalytics.objects.filter(qr_code=self.qr).count(), 1)

    def test_gps_denial_does_not_create_second_event(self):
        self.qr.require_gps = True
        self.qr.save(update_fields=['require_gps'])
        with patch('dynamic_qr.views.render', return_value=HttpResponse('')):
            self.client.get(f'/qr/r/{self.qr.short_code}/')
            response = self.client.post(
                f'/qr/r/{self.qr.short_code}/', {'gps_denied': 'true'}
            )
        self.assertEqual(response.status_code, 200)
        event = QRAnalytics.objects.get(qr_code=self.qr)
        self.assertEqual(QRAnalytics.objects.filter(qr_code=self.qr).count(), 1)
        self.assertEqual(event.gps_permission, 'denied')
        self.assertEqual(event.gps_latitude, None)

    def test_private_address_detection_and_gps_validation(self):
        self.assertTrue(utils.is_private_address('127.0.0.1'))
        self.assertTrue(utils.is_private_address('::1'))
        self.assertTrue(utils.is_private_address('192.168.1.10'))
        self.assertFalse(utils.is_private_address('8.8.8.8'))

        self.qr.require_gps = True
        self.qr.save(update_fields=['require_gps'])
        with patch('dynamic_qr.views.render', return_value=HttpResponse('')):
            self.client.get(f'/qr/r/{self.qr.short_code}/')
            self.client.post(
                f'/qr/r/{self.qr.short_code}/',
                {'gps_lat': '91', 'gps_lon': '0', 'gps_accuracy': '1'},
            )
        event = QRAnalytics.objects.get(qr_code=self.qr)
        self.assertEqual(event.gps_permission, 'pending')
        self.assertIsNone(event.gps_latitude)

    @patch.object(utils.QRAnalytics.objects, 'create')
    def test_event_failure_does_not_increment_cached_counter(self, create_event):
        create_event.side_effect = RuntimeError('database unavailable')

        self.client.get(
            f'/qr/r/{self.qr.short_code}/',
            HTTP_USER_AGENT=self.user_agent,
        )

        self.qr.refresh_from_db()
        self.assertEqual(self.qr.scan_count, 0)
        self.assertEqual(QRAnalytics.objects.filter(qr_code=self.qr).count(), 0)

    def test_gps_save_twice_updates_one_visit(self):
        self.qr.require_gps = True
        self.qr.save(update_fields=['require_gps'])

        with patch('dynamic_qr.views.render', return_value=HttpResponse('')):
            self.client.get(f'/qr/r/{self.qr.short_code}/')

        response = self.client.post(
            f'/qr/r/{self.qr.short_code}/',
            data=json.dumps({
                'latitude': 13.0827,
                'longitude': 80.2707,
                'accuracy': 8,
                'permission': 'granted',
            }),
            content_type='application/json',
        )
        self.assertEqual(response.status_code, 200)
        duplicate_response = self.client.post(
            f'/qr/r/{self.qr.short_code}/',
            data=json.dumps({
                'latitude': 13.0827,
                'longitude': 80.2707,
                'accuracy': 8,
                'permission': 'granted',
            }),
            content_type='application/json',
        )
        self.assertEqual(duplicate_response.status_code, 302)
        self.assertEqual(QRAnalytics.objects.filter(qr_code=self.qr).count(), 1)
        self.qr.refresh_from_db()
        self.assertEqual(self.qr.scan_count, 1)

    def test_analytics_counts_canonical_visits_and_gps_states(self):
        now = timezone.now()
        QRAnalytics.objects.create(
            qr_code=self.qr,
            timestamp=now,
            country='India',
            country_code='IN',
            city='Chennai',
            location_source='gps',
            gps_permission='granted',
            gps_latitude=13.08,
            gps_longitude=80.27,
            source='direct',
            visitor_id='visitor-one',
            redirect_result='redirect_success',
        )
        QRAnalytics.objects.create(
            qr_code=self.qr,
            timestamp=now,
            country='India',
            country_code='IN',
            city='Chennai',
            location_source='ip',
            gps_permission='denied',
            source='direct',
            visitor_id='visitor-two',
            redirect_result='gps_denied',
        )

        self.client.force_login(self.user)
        session = self.client.session
        session['is_dqr_user'] = True
        session.save()
        response = self.client.get(f'/qr/short-url/analytics/{self.qr.id}/?range=7days')

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.context['total_clicks'], 2)
        self.assertEqual(response.context['unique_clicks'], 2)
        summary = response.context['perf_summary']
        self.assertEqual(summary['gps_requests'], 2)
        self.assertEqual(summary['gps_granted'], 1)
        self.assertEqual(summary['gps_denied'], 1)
        self.assertEqual(summary['capture_rate'], 50)
        self.assertEqual(sum(row['count'] for row in response.context['ts_stats']), 2)
        self.assertEqual(response.context['country_stats'][0]['count'], 2)
        self.assertEqual(response.context['country_stats'][0]['percentage'], 100)

    def test_six_canonical_visits_stay_six_across_analytics(self):
        for index in range(6):
            QRAnalytics.objects.create(
                qr_code=self.qr,
                country='India',
                country_code='IN',
                city='Chennai',
                location_source='ip',
                gps_permission='not_required',
                source='qr' if index < 2 else 'direct',
                is_qr_scan=index < 2,
                visitor_id=f'visitor-{index}',
                redirect_result='redirect_success',
            )

        self.client.force_login(self.user)
        session = self.client.session
        session['is_dqr_user'] = True
        session.save()
        response = self.client.get(f'/qr/short-url/analytics/{self.qr.id}/?range=7days')

        self.assertEqual(response.context['total_clicks'], 6)
        self.assertEqual(response.context['page_obj'].paginator.count, 6)
        self.assertEqual(sum(row['count'] for row in response.context['ts_stats']), 6)
        self.assertEqual(response.context['country_stats'][0]['percentage'], 100)
        self.assertEqual(response.context['perf_summary']['top_country_pct'], 100)
