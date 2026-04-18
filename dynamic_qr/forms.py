from django import forms
from django.contrib.auth.models import User
from django.contrib.auth.forms import AuthenticationForm
from django.contrib.auth.password_validation import validate_password


class DynamicQRLoginForm(AuthenticationForm):
    """Login form used exclusively for the Dynamic QR feature."""
    username = forms.CharField(
        widget=forms.TextInput(attrs={
            'class': 'dqr-field',
            'placeholder': 'Username or Email',
            'autocomplete': 'username',
        })
    )
    password = forms.CharField(
        widget=forms.PasswordInput(attrs={
            'class': 'dqr-field',
            'placeholder': '••••••••',
            'autocomplete': 'current-password',
        })
    )


class DynamicQRRegisterForm(forms.ModelForm):
    """Registration form used exclusively for the Dynamic QR feature."""
    username = forms.CharField(
        max_length=150,
        widget=forms.TextInput(attrs={
            'class': 'dqr-field',
            'placeholder': 'Choose a username',
        })
    )
    email = forms.EmailField(
        widget=forms.EmailInput(attrs={
            'class': 'dqr-field',
            'placeholder': 'your@email.com',
        })
    )
    password1 = forms.CharField(
        label='Password',
        widget=forms.PasswordInput(attrs={
            'class': 'dqr-field',
            'placeholder': 'Create a password',
        })
    )
    password2 = forms.CharField(
        label='Confirm Password',
        widget=forms.PasswordInput(attrs={
            'class': 'dqr-field',
            'placeholder': 'Confirm password',
        })
    )

    class Meta:
        model = User
        fields = ['username', 'email']

    def clean_email(self):
        email = self.cleaned_data.get('email')
        if User.objects.filter(email=email).exists():
            raise forms.ValidationError("An account with this email already exists.")
        return email

    def clean_username(self):
        username = self.cleaned_data.get('username')
        if User.objects.filter(username=username).exists():
            raise forms.ValidationError("This username is already taken.")
        return username

    def clean(self):
        cleaned_data = super().clean()
        p1 = cleaned_data.get('password1')
        p2 = cleaned_data.get('password2')
        if p1 and p2 and p1 != p2:
            self.add_error('password2', 'Passwords do not match.')
        if p1:
            try:
                validate_password(p1)
            except forms.ValidationError as e:
                self.add_error('password1', e)
        return cleaned_data

    def save(self, commit=True):
        user = super().save(commit=False)
        user.set_password(self.cleaned_data['password1'])
        if commit:
            user.save()
        return user


class ForgotPasswordForm(forms.Form):
    """Forgot password — enter email to receive OTP."""
    email = forms.EmailField(
        widget=forms.EmailInput(attrs={
            'class': 'dqr-field',
            'placeholder': 'Enter your registered email',
        })
    )


class OTPVerifyForm(forms.Form):
    """OTP verification form."""
    otp = forms.CharField(
        max_length=6,
        min_length=6,
        widget=forms.TextInput(attrs={
            'class': 'dqr-field text-center tracking-[0.5em] text-2xl font-black',
            'placeholder': '______',
            'maxlength': '6',
            'autocomplete': 'one-time-code',
        })
    )


class ResetPasswordForm(forms.Form):
    """Reset password after OTP verification."""
    new_password = forms.CharField(
        widget=forms.PasswordInput(attrs={
            'class': 'dqr-field',
            'placeholder': 'New password',
        })
    )
    confirm_password = forms.CharField(
        widget=forms.PasswordInput(attrs={
            'class': 'dqr-field',
            'placeholder': 'Confirm new password',
        })
    )

    def clean(self):
        cleaned_data = super().clean()
        p1 = cleaned_data.get('new_password')
        p2 = cleaned_data.get('confirm_password')
        if p1 and p2 and p1 != p2:
            self.add_error('confirm_password', 'Passwords do not match.')
        if p1:
            try:
                validate_password(p1)
            except forms.ValidationError as e:
                self.add_error('new_password', e)
        return cleaned_data


class DynamicQRForm(forms.Form):
    """Form for creating/editing a dynamic QR code."""
    qr_name = forms.CharField(
        max_length=200,
        widget=forms.TextInput(attrs={
            'class': 'dqr-field',
            'placeholder': 'My Website QR',
        })
    )
    destination_url = forms.URLField(
        max_length=2000,
        widget=forms.URLInput(attrs={
            'class': 'dqr-field',
            'placeholder': 'https://example.com',
        })
    )
    fg_color = forms.CharField(max_length=10, initial='#000000', required=False)
    bg_color = forms.CharField(max_length=10, initial='#ffffff', required=False)
    body_style = forms.CharField(max_length=20, initial='square', required=False)
    eye_style = forms.CharField(max_length=20, initial='square', required=False)
    ball_style = forms.CharField(max_length=20, initial='square', required=False)
