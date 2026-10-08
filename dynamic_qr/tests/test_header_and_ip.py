from django.test import TestCase, RequestFactory
from django.conf import settings
from dynamic_qr.utils import get_client_ip

class TrustedIPTests(TestCase):
    def setUp(self):
        self.factory = RequestFactory()

    def test_no_trusted_proxy(self):
        # By default, TRUSTED_PROXY_IPS is empty or not set
        request = self.factory.get('/', REMOTE_ADDR='1.2.3.4', HTTP_X_FORWARDED_FOR='8.8.8.8')
        self.assertEqual(get_client_ip(request), '1.2.3.4')

    def test_trusted_proxy_configuration(self):
        with self.settings(TRUSTED_PROXY_IPS=['10.0.0.1']):
            # Proxy is trusted, should extract client IP from XFF
            request = self.factory.get('/', REMOTE_ADDR='10.0.0.1', HTTP_X_FORWARDED_FOR='8.8.8.8')
            self.assertEqual(get_client_ip(request), '8.8.8.8')

    def test_multiple_trusted_proxies(self):
        with self.settings(TRUSTED_PROXY_IPS=['10.0.0.1', '10.0.0.2']):
            # client=8.8.8.8, proxy1=10.0.0.1, proxy2=10.0.0.2 (REMOTE_ADDR)
            request = self.factory.get('/', REMOTE_ADDR='10.0.0.2', HTTP_X_FORWARDED_FOR='8.8.8.8, 10.0.0.1')
            self.assertEqual(get_client_ip(request), '8.8.8.8')

    def test_spoof_attempt(self):
        with self.settings(TRUSTED_PROXY_IPS=['10.0.0.1']):
            # Attacker sends fake XFF directly
            request = self.factory.get('/', REMOTE_ADDR='attacker.ip', HTTP_X_FORWARDED_FOR='fake.ip')
            self.assertEqual(get_client_ip(request), 'attacker.ip')

    def test_malformed_xff(self):
        with self.settings(TRUSTED_PROXY_IPS=['10.0.0.1']):
            request = self.factory.get('/', REMOTE_ADDR='10.0.0.1', HTTP_X_FORWARDED_FOR='not_an_ip, 8.8.8.8')
            self.assertEqual(get_client_ip(request), '8.8.8.8')
            
    def test_ipv6(self):
        with self.settings(TRUSTED_PROXY_IPS=['2001:db8::1']):
            request = self.factory.get('/', REMOTE_ADDR='2001:db8::1', HTTP_X_FORWARDED_FOR='2001:db8::2')
            self.assertEqual(get_client_ip(request), '2001:db8::2')

class HeaderConsistencyTests(TestCase):
    def test_header_validation_max_length(self):
        from dynamic_qr.validators import validate_header
        from django.core.exceptions import ValidationError
        
        # Valid 20 char
        validate_header('a' * 20)
        
        # Invalid 21 char
        with self.assertRaises(ValidationError):
            validate_header('a' * 21)
            
    def test_header_validation_characters(self):
        from dynamic_qr.validators import validate_header
        from django.core.exceptions import ValidationError
        
        # Invalid chars
        with self.assertRaises(ValidationError):
            validate_header('invalid/slash')
        with self.assertRaises(ValidationError):
            validate_header('invalid?query')
