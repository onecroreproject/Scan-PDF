from django import forms
from converter.models import HeroVideo

class HeroVideoForm(forms.ModelForm):
    class Meta:
        model = HeroVideo
        fields = ['section', 'title', 'video', 'order', 'is_active']
        widgets = {
            'section': forms.Select(attrs={
                'class': 'w-full px-4 py-2 border border-surface-300 rounded-xl focus:ring-2 focus:ring-brand-500 focus:border-brand-500 transition-all bg-white'
            }),
            'title': forms.TextInput(attrs={
                'class': 'w-full px-4 py-2 border border-surface-300 rounded-xl focus:ring-2 focus:ring-brand-500 focus:border-brand-500 transition-all',
                'placeholder': 'E.g., Summer Promo (Optional)'
            }),
            'video': forms.ClearableFileInput(attrs={
                'class': 'w-full px-4 py-2 border border-surface-300 rounded-xl focus:ring-2 focus:ring-brand-500 focus:border-brand-500 transition-all bg-white file:mr-4 file:py-2 file:px-4 file:rounded-full file:border-0 file:text-sm file:font-semibold file:bg-brand-50 file:text-brand-700 hover:file:bg-brand-100'
            }),
            'order': forms.NumberInput(attrs={
                'class': 'w-full px-4 py-2 border border-surface-300 rounded-xl focus:ring-2 focus:ring-brand-500 focus:border-brand-500 transition-all',
                'min': '0'
            }),
            'is_active': forms.CheckboxInput(attrs={
                'class': 'w-5 h-5 text-brand-600 border-surface-300 rounded focus:ring-brand-500 cursor-pointer'
            })
        }
