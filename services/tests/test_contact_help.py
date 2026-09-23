from django.test import TestCase
from django.urls import reverse


class ContactAndHelpPagesTest(TestCase):
    def test_help_page_loads(self):
        response = self.client.get(reverse('services:help'))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, 'How Can We Help?')

    def test_contact_page_loads(self):
        response = self.client.get(reverse('services:contact'))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, 'Talk')
        self.assertContains(response, 'Contact Us')

    def test_contact_form_rejects_blank_message(self):
        response = self.client.post(reverse('services:contact'), {
            'name': 'John Doe',
            'email': 'john@example.com',
            'phone': '+91 9988776655',
            'subject': 'Test',
            'category': 'General Enquiry',
            'message': '   ',
            'website': '',
        })
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, 'message')

    def test_contact_ajax_submit_returns_json(self):
        response = self.client.post(
            reverse('services:contact'),
            {
                'name': 'John Doe',
                'email': 'john@example.com',
                'phone': '+91 9988776655',
                'subject': 'Test enquiry',
                'category': 'General Enquiry',
                'message': 'This is a valid support enquiry message.',
                'website': '',
            },
            HTTP_X_REQUESTED_WITH='XMLHttpRequest',
        )
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json()['success'], True)
        self.assertIn('ticket_id', response.json())

    def test_help_ajax_validation_error_returns_json(self):
        response = self.client.post(
            reverse('services:help'),
            {
                'name': 'Jane Doe',
                'email': 'invalid-email',
                'subject': 'Help me',
                'category': 'General Help',
                'message': 'Short',
                'website': '',
            },
            HTTP_X_REQUESTED_WITH='XMLHttpRequest',
        )
        self.assertEqual(response.status_code, 400)
        self.assertEqual(response.json()['success'], False)
        self.assertIn('validation_error', response.json()['status'])
