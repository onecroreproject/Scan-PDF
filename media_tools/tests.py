from django.test import TestCase, Client
from django.urls import reverse
from media_tools.services.downloader_service import is_supported_url

class VideoDownloaderTests(TestCase):
    def setUp(self):
        self.client = Client()
        
    def test_downloader_pages_render(self):
        platforms = [
            'instagram', 
            'twitter', 
            'facebook', 
            'threads', 
            'youtube'
        ]
        
        for platform in platforms:
            url_name = f'media_tools:downloader_{platform}'
            response = self.client.get(reverse(url_name))
            self.assertEqual(response.status_code, 200, f"{platform} page failed to load")
            self.assertContains(response, 'Download', msg_prefix=f"{platform} template missing expected content")
            
    def test_url_validation_success(self):
        valid_urls = [
            'https://www.youtube.com/watch?v=dQw4w9WgXcQ',
            'https://youtu.be/dQw4w9WgXcQ',
            'https://www.instagram.com/p/abcdefg/',
            'https://twitter.com/i/status/123456789',
            'https://x.com/i/status/123456789',
            'https://www.facebook.com/watch/?v=123456789',
            'https://www.threads.net/@user/post/abcdefg',
            'https://www.threads.com/share/Qnvtb2zwF/',
            'https://threads.com/share/Qnvtb2zwF/'
        ]
        for url in valid_urls:
            self.assertIsNotNone(is_supported_url(url), f"Failed to validate supported URL: {url}")
            if 'threads' in url:
                self.assertEqual(is_supported_url(url), 'threads')
            
    def test_url_validation_rejection(self):
        invalid_urls = [
            'http://localhost:8000/media',
            'http://127.0.0.1/video',
            'https://my-instagram.com/post',
            'https://evil-youtube.com/watch',
            'file:///etc/passwd',
            'ftp://example.com/file',
            'https://youtube.com@localhost/',
            'https://instagram.com@192.168.1.1/',
            'https://evilthreads.com/test',
            'https://threads.com.evil.com/test',
            'https://notthreads.com/test',
            'https://threads-example.com/test',
            'http://threads.com@127.0.0.1/',
            'http://threads.com@localhost/',
            'file://threads.com/test'
        ]
        for url in invalid_urls:
            self.assertIsNone(is_supported_url(url), f"Failed to reject unsupported URL: {url}")
