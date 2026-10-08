from django.test import TestCase, RequestFactory
from django.conf import settings
from unittest.mock import patch

class RateLimitingTests(TestCase):
    def setUp(self):
        self.factory = RequestFactory()

    @patch('dynamic_qr.security.cache.incr')
    @patch('dynamic_qr.security.cache.set')
    def test_password_limiter_increment(self, mock_set, mock_incr):
        from dynamic_qr.security import PasswordAttemptLimiter
        
        # Simulate normal failure
        mock_incr.return_value = 1
        PasswordAttemptLimiter.record_failure(1, 'client_a')
        mock_incr.assert_called_once()
        
    @patch('dynamic_qr.security.cache.incr')
    @patch('dynamic_qr.security.cache.set')
    def test_password_limiter_lockout(self, mock_set, mock_incr):
        from dynamic_qr.security import PasswordAttemptLimiter
        
        # Simulate 5th failure
        mock_incr.return_value = 5
        PasswordAttemptLimiter.record_failure(1, 'client_a')
        
        # Assert lock key is set
        self.assertEqual(mock_set.call_args[0][0], PasswordAttemptLimiter.get_lock_key(1, 'client_a'))
        self.assertEqual(mock_set.call_args[0][1], '1')

    @patch('dynamic_qr.security.cache.incr')
    def test_redirect_abuse_limiter(self, mock_incr):
        from dynamic_qr.security import RedirectAbuseLimiter
        
        # Under limit
        mock_incr.return_value = 100
        self.assertFalse(RedirectAbuseLimiter.is_rate_limited('client_b'))
        
        # Over limit
        mock_incr.return_value = 601
        self.assertTrue(RedirectAbuseLimiter.is_rate_limited('client_b'))

    @patch('dynamic_qr.security.cache.get')
    def test_cache_failure_fail_closed_password(self, mock_get):
        from dynamic_qr.security import PasswordAttemptLimiter
        
        # Exception during cache read
        mock_get.side_effect = Exception("Redis Down")
        
        # Fails closed
        self.assertTrue(PasswordAttemptLimiter.is_locked(1, 'client_c'))

    @patch('dynamic_qr.security.cache.incr')
    def test_cache_failure_fail_open_redirect(self, mock_incr):
        from dynamic_qr.security import RedirectAbuseLimiter
        
        # Exception during cache read
        mock_incr.side_effect = Exception("Redis Down")
        
        # Fails open to preserve availability
        self.assertFalse(RedirectAbuseLimiter.is_rate_limited('client_d'))
