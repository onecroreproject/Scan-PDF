from django import forms

from .models import ContactEnquiry

HELP_CATEGORIES = [
    'General Help', 'PDF Tools', 'Image Tools', 'Video Tools', 'QR Code', 'Short URL',
    'Account Issue', 'Login Issue', 'Pricing', 'Billing', 'Technical Issue', 'Bug Report', 'Other'
]

CONTACT_CATEGORIES = [
    'General Enquiry', 'Technical Support', 'Account Issue', 'Pricing', 'Billing',
    'Business Enquiry', 'Partnership', 'Bug Report', 'Feedback', 'Other'
]


class ContactEnquiryForm(forms.ModelForm):
    website = forms.CharField(required=False, widget=forms.HiddenInput())
    source = forms.ChoiceField(choices=ContactEnquiry.SOURCE_CHOICES, required=False, widget=forms.HiddenInput())
    name = forms.CharField(max_length=120, strip=True)
    email = forms.EmailField(max_length=255)
    phone = forms.CharField(max_length=30, required=False, strip=True)
    category = forms.ChoiceField(required=True)
    subject = forms.CharField(max_length=200, strip=True)
    message = forms.CharField(widget=forms.Textarea(attrs={'rows': 6, 'placeholder': 'Tell us how we can help...'}), strip=True)

    class Meta:
        model = ContactEnquiry
        fields = ['name', 'email', 'phone', 'category', 'subject', 'message', 'source']

    def __init__(self, *args, **kwargs):
        source = kwargs.pop('source', ContactEnquiry.SOURCE_CONTACT)
        super().__init__(*args, **kwargs)
        source = source or ContactEnquiry.SOURCE_CONTACT
        self.fields['source'].initial = source
        self.fields['name'].widget.attrs.update({
            'class': 'scan-contact-form-control',
            'placeholder': 'Enter your full name',
        })
        self.fields['email'].widget.attrs.update({
            'class': 'scan-contact-form-control',
            'placeholder': 'you@example.com',
        })
        self.fields['phone'].widget.attrs.update({
            'class': 'scan-contact-form-control',
            'placeholder': 'Phone number',
        })
        self.fields['subject'].widget.attrs.update({
            'class': 'scan-contact-form-control',
            'placeholder': 'Enter subject',
        })
        self.fields['category'].widget.attrs.update({
            'class': 'scan-contact-form-control',
        })
        self.fields['message'].widget.attrs.update({
            'class': 'scan-contact-form-control',
            'placeholder': 'Tell us how we can help...',
            'rows': 6,
        })
        category_values = HELP_CATEGORIES if source == ContactEnquiry.SOURCE_HELP else CONTACT_CATEGORIES
        self.fields['category'].choices = [('', 'Select a category')] + [(value, value) for value in category_values]

    def clean_name(self):
        name = (self.cleaned_data.get('name') or '').strip()
        if not name:
            raise forms.ValidationError('Please enter your name.')
        if len(name) > 120:
            raise forms.ValidationError('Name must be 120 characters or less.')
        return name

    def clean_subject(self):
        subject = (self.cleaned_data.get('subject') or '').strip()
        if not subject:
            raise forms.ValidationError('Please enter a subject.')
        if len(subject) > 200:
            raise forms.ValidationError('Subject must be 200 characters or less.')
        return subject

    def clean_message(self):
        message = (self.cleaned_data.get('message') or '').strip()
        if not message:
            raise forms.ValidationError('Please enter your message.')
        if len(message) < 10:
            raise forms.ValidationError('Message must be at least 10 characters long.')
        if len(message) > 5000:
            raise forms.ValidationError('Message must be 5000 characters or less.')
        return message

    def clean_website(self):
        if self.cleaned_data.get('website'):
            raise forms.ValidationError('Spam submission rejected.')
        return ''

    def clean(self):
        cleaned = super().clean()
        if cleaned.get('website'):
            raise forms.ValidationError('Spam submission rejected.')
        if cleaned.get('message'):
            cleaned['message'] = cleaned['message'].strip()
        if cleaned.get('name'):
            cleaned['name'] = cleaned['name'].strip()
        if cleaned.get('subject'):
            cleaned['subject'] = cleaned['subject'].strip()
        return cleaned
