import unittest
from unittest.mock import patch, MagicMock
from django.test import RequestFactory, TestCase
import json

from video_downloader.services import _categorize_error, YTDLPError
from video_downloader.views import analyze_url

class TestYouTubeDownloader(TestCase):
    def setUp(self):
        self.factory = RequestFactory()

    def test_bot_challenge_categorization(self):
        code, msg = _categorize_error(Exception("Sign in to confirm you're not a bot"), "https://youtube.com/watch?v=123")
        self.assertEqual(code, "YOUTUBE_BOT_CHALLENGE")
        
    def test_unavailable_video_categorization(self):
        code, msg = _categorize_error(Exception("Video unavailable"), "https://youtube.com/watch?v=123")
        self.assertEqual(code, "VIDEO_UNAVAILABLE")
        
    def test_rate_limit_categorization(self):
        code, msg = _categorize_error(Exception("HTTP Error 429: Too Many Requests"), "https://youtube.com/watch?v=123")
        self.assertEqual(code, "RATE_LIMITED")

    @patch('video_downloader.services.analyze_video')
    def test_analyze_url_catches_ytdlp_error(self, mock_analyze):
        mock_analyze.side_effect = YTDLPError("YOUTUBE_BOT_CHALLENGE", "Safe sanitized message.")
        
        request = self.factory.post(
            '/api/analyze/',
            data=json.dumps({"url": "https://www.youtube.com/watch?v=123"}),
            content_type='application/json'
        )
        
        response = analyze_url(request)
        self.assertEqual(response.status_code, 400)
        
        data = json.loads(response.content)
        self.assertFalse(data['success'])
        self.assertEqual(data['error_code'], "YOUTUBE_BOT_CHALLENGE")
        self.assertEqual(data['message'], "Safe sanitized message.")
