from django.contrib import admin
from .models import DynamicQRCode, OTPVerification


@admin.register(DynamicQRCode)
class DynamicQRCodeAdmin(admin.ModelAdmin):
    list_display = ('qr_name', 'short_code', 'user', 'destination_url', 'scan_count', 'is_active', 'created_at')
    list_filter = ('is_active', 'created_at', 'qr_type')
    search_fields = ('qr_name', 'short_code', 'destination_url', 'user__username')
    readonly_fields = ('id', 'short_code', 'scan_count', 'created_at', 'updated_at')


@admin.register(OTPVerification)
class OTPVerificationAdmin(admin.ModelAdmin):
    list_display = ('email', 'otp_code', 'is_used', 'attempts', 'created_at')
    list_filter = ('is_used', 'created_at')
    search_fields = ('email',)
