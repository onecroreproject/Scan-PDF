from django.test import TestCase, Client
from django.urls import reverse
from django.contrib.auth.models import User
from dynamic_qr.models import DynamicQRCode
import json
from unittest.mock import patch

class ShortUrlHardeningTests(TestCase):
    def setUp(self):
        self.user = User.objects.create_user(username='testuser', password='testpassword')
        self.client = Client()
        self.client.login(username='testuser', password='testpassword')
        
    def test_ssrf_protection(self):
        """Test that SSRF attempts to localhost/private IPs are blocked during creation."""
        response = self.client.post(reverse('dynamic_qr:create'), {
            'destination_url': 'http://127.0.0.1/admin',
            'qr_name': 'SSRF Attempt',
            'qr_type': 'custom-url'
        })
        self.assertEqual(response.status_code, 400)
        self.assertIn('Private IP', response.json().get('error', ''))
        
    def test_protocol_validation(self):
        """Test that only whitelisted protocols are allowed."""
        response = self.client.post(reverse('dynamic_qr:create'), {
            'destination_url': 'javascript:alert(1)',
            'qr_name': 'XSS Attempt',
            'qr_type': 'custom-url'
        })
        self.assertEqual(response.status_code, 400)
        self.assertIn('not allowed', response.json().get('error', ''))
        
    def test_short_code_collision_retry(self):
        """Test that collision in short code generation is handled via retry."""
        qr1 = DynamicQRCode.objects.create(user=self.user, short_code='ABCDEFGH')
        
        with patch('dynamic_qr.models.generate_short_code') as mock_generate:
            # First attempt returns existing, second returns new
            mock_generate.side_effect = ['ABCDEFGH', 'HGFEDCBA']
            
            response = self.client.post(reverse('dynamic_qr:create'), {
                'destination_url': 'https://google.com',
                'qr_name': 'Retry Test',
                'qr_type': 'custom-url'
            })
            self.assertEqual(response.status_code, 200)
            self.assertEqual(mock_generate.call_count, 2)
            
    def test_password_brute_force_protection(self):
        """Test that password brute force is prevented via cache lockout."""
        qr = DynamicQRCode.objects.create(
            user=self.user, 
            short_code='SECURE', 
            password='hashedpassword123'  # Mock hash
        )
        url = reverse('dynamic_qr:redirect', args=['SECURE'])
        
        # 5 failed attempts
        for _ in range(5):
            response = self.client.post(url, {'password': 'wrong'})
            
        # 6th attempt should return 429/lockout message
        response = self.client.post(url, {'password': 'wrong'})
        self.assertContains(response, 'Too many failed attempts', status_code=200)
