from django import forms


class FileUploadForm(forms.Form):
    """Form for handling file uploads for conversion."""
    file = forms.FileField(
        label='Select a file',
        help_text='Upload a file to convert',
        widget=forms.ClearableFileInput(attrs={
            'class': 'hidden',
            'id': 'file-input',
        })
    )
