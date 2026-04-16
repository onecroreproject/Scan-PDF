"""
Views for Dynamic QR Code feature.
This module handles:
- Dedicated login/register (only for dynamic QR, not for rest of project)
- Forgot password with Gmail OTP verification
- Dynamic QR dashboard (list, create, edit, delete)
- QR redirect endpoint (short code → destination)
"""
import random
import json
import os
from datetime import timedelta

from django.shortcuts import render, redirect, get_object_or_404
from django.http import JsonResponse, HttpResponseRedirect
from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.models import User
from django.contrib.auth.decorators import login_required
from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.http import require_POST
from django.utils import timezone
from django.core.mail import send_mail
from django.conf import settings

from .models import DynamicQRCode, OTPVerification
from .forms import (
    DynamicQRLoginForm,
    DynamicQRRegisterForm,
    ForgotPasswordForm,
    OTPVerifyForm,
    ResetPasswordForm,
    DynamicQRForm,
)


# ═══════════════════════════════════════════════════════════════
# HELPER: Check if user is logged in for Dynamic QR
# ═══════════════════════════════════════════════════════════════
def dqr_login_required(view_func):
    """Decorator: redirect to dynamic QR login if not authenticated or not a QR user."""
    def wrapper(request, *args, **kwargs):
        # Isolation: Check if authenticated AND has the dqr flag
        if not request.user.is_authenticated or not request.session.get('is_dqr_user'):
            return redirect('dynamic_qr:login')
        return view_func(request, *args, **kwargs)
    return wrapper


from django.db import connection

def dqr_repair_db(request):
    """Utility view to manually add missing columns/tables to SQLite via browser."""
    # Allow superusers OR regular authenticated users for this specific repair
    if not request.user.is_authenticated:
        from django.http import HttpResponse
        return HttpResponse("Unauthorized.", status=403)
    
    from django.http import HttpResponse
    with connection.cursor() as cursor:
        results = []
        
        # 0. Ensure Main DynamicQRCode Table exists
        try:
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS "dynamic_qr_dynamicqrcode" (
                    "id" uuid NOT NULL PRIMARY KEY,
                    "short_code" varchar(20) NOT NULL UNIQUE,
                    "qr_name" varchar(200) NOT NULL,
                    "destination_url" varchar(2000) NULL,
                    "qr_data" json NOT NULL,
                    "qr_type" varchar(40) NOT NULL DEFAULT 'url',
                    "fg_color" varchar(10) NOT NULL DEFAULT '#000000',
                    "bg_color" varchar(10) NOT NULL DEFAULT '#ffffff',
                    "body_style" varchar(20) NOT NULL DEFAULT 'square',
                    "eye_style" varchar(20) NOT NULL DEFAULT 'square',
                    "ball_style" varchar(20) NOT NULL DEFAULT 'square',
                    "logo" varchar(100) NULL,
                    "scan_count" integer unsigned NOT NULL DEFAULT 0,
                    "is_active" bool NOT NULL DEFAULT 1,
                    "created_at" datetime NOT NULL,
                    "updated_at" datetime NOT NULL,
                    "user_id" integer NOT NULL REFERENCES "auth_user" ("id") DEFERRABLE INITIALLY DEFERRED
                );
            """)
            results.append("✅ Main table 'dynamic_qr_dynamicqrcode' is ready.")
        except Exception as e:
            results.append(f"❌ Error with main table: {str(e)}")

        # 1. Add missing columns to DynamicQRCode (for existing users)
        cols = [
            ("qr_data", "JSON"),
            ("qr_type", "VARCHAR(40) DEFAULT 'url'"),
            ("logo", "VARCHAR(100) NULL"),
            ("body_style", "VARCHAR(20) DEFAULT 'square'"),
            ("is_active", "BOOLEAN DEFAULT 1"),
            ("file_content", "VARCHAR(100) NULL"),
            ("eye_style", "VARCHAR(20) DEFAULT 'square'"),
            ("ball_style", "VARCHAR(20) DEFAULT 'square'")
        ]
        for col_name, col_type in cols:
            try:
                cursor.execute(f"ALTER TABLE dynamic_qr_dynamicqrcode ADD COLUMN {col_name} {col_type};")
                results.append(f"✅ Added column: {col_name}")
            except Exception as e:
                results.append(f"ℹ️ Column '{col_name}' already exists.")

        # 2. Create the Analytics Table
        try:
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS "dynamic_qr_qranalytics" (
                    "id" integer NOT NULL PRIMARY KEY AUTOINCREMENT,
                    "timestamp" datetime NOT NULL,
                    "ip_address" char(39) NULL,
                    "user_agent" text NULL,
                    "browser" varchar(50) NULL,
                    "os" varchar(50) NULL,
                    "device_type" varchar(50) NULL,
                    "country" varchar(100) NOT NULL DEFAULT 'Unknown',
                    "city" varchar(100) NOT NULL DEFAULT 'Unknown',
                    "qr_code_id" uuid NOT NULL REFERENCES "dynamic_qr_dynamicqrcode" ("id") DEFERRABLE INITIALLY DEFERRED
                );
            """)
            cursor.execute('CREATE INDEX IF NOT EXISTS "dynamic_qr_analytics_qr_id" ON "dynamic_qr_qranalytics" ("qr_code_id");')
            results.append("✅ Table 'dynamic_qr_qranalytics' is ready.")
        except Exception as e:
            results.append(f"❌ Error with analytics table: {str(e)}")

        # 3. Create OTP Table
        try:
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS "dynamic_qr_otpverification" (
                    "id" integer NOT NULL PRIMARY KEY AUTOINCREMENT,
                    "email" varchar(254) NOT NULL,
                    "otp_code" varchar(6) NOT NULL,
                    "created_at" datetime NOT NULL,
                    "is_used" bool NOT NULL DEFAULT 0,
                    "attempts" integer unsigned NOT NULL DEFAULT 0
                );
            """)
            results.append("✅ Table 'dynamic_qr_otpverification' is ready.")
        except Exception as e:
            results.append(f"❌ Error with OTP table: {str(e)}")
    
    return HttpResponse("<h3>Database Repair Results</h3>" + "<br>".join(results) + "<br><br><b>All fixed.</b> <a href='/qr/dashboard/'>Return to Dashboard</a>")

# ═══════════════════════════════════════════════════════════════
# AUTH: LOGIN
# ═══════════════════════════════════════════════════════════════
def dqr_login_view(request):
    """Login page for dynamic QR feature only."""
    # Only redirect if they are fully authenticated for the QR system
    if request.user.is_authenticated and request.session.get('is_dqr_user'):
        return redirect('dynamic_qr:dashboard')

    error = None
    if request.method == 'POST':
        username = request.POST.get('username', '').strip()
        password = request.POST.get('password', '')

        # Allow login with email or username
        user = None
        if '@' in username:
            try:
                user_obj = User.objects.get(email=username)
                user = authenticate(request, username=user_obj.username, password=password)
            except User.DoesNotExist:
                user = None
        else:
            user = authenticate(request, username=username, password=password)

        if user is not None:
            login(request, user)
            # Mark this session as a Dynamic QR session for isolation
            request.session['is_dqr_user'] = True
            next_url = request.GET.get('next', '')
            return redirect(next_url if next_url else 'dynamic_qr:dashboard')
        else:
            error = "Invalid username/email or password."

    return render(request, 'dynamic_qr/login.html', {'error': error})


# ═══════════════════════════════════════════════════════════════
# AUTH: REGISTER
# ═══════════════════════════════════════════════════════════════
def dqr_register_view(request):
    """Register page for dynamic QR feature only."""
    if request.user.is_authenticated and request.session.get('is_dqr_user'):
        return redirect('dynamic_qr:dashboard')

    form = DynamicQRRegisterForm()
    if request.method == 'POST':
        form = DynamicQRRegisterForm(request.POST)
        if form.is_valid():
            user = form.save()
            login(request, user)
            request.session['is_dqr_user'] = True
            return redirect('dynamic_qr:dashboard')

    return render(request, 'dynamic_qr/register.html', {'form': form})


# ═══════════════════════════════════════════════════════════════
# AUTH: LOGOUT
# ═══════════════════════════════════════════════════════════════
def dqr_logout_view(request):
    """Logout from dynamic QR session."""
    logout(request)
    # Ensure session is completely flushed to avoid any isolation leakage
    request.session.flush()
    return redirect('dynamic_qr:login')


# ═══════════════════════════════════════════════════════════════
# AUTH: FORGOT PASSWORD — Send OTP via Gmail
# ═══════════════════════════════════════════════════════════════
def dqr_forgot_password_view(request):
    """Forgot password: enter email → receive OTP via Gmail."""
    message = None
    error = None

    if request.method == 'POST':
        email = request.POST.get('email', '').strip()

        if not email:
            error = "Please enter your email address."
        elif not User.objects.filter(email=email).exists():
            error = "No account found with this email address."
        else:
            # Generate 6-digit OTP
            otp = ''.join([str(random.randint(0, 9)) for _ in range(6)])

            # Invalidate previous OTPs for this email
            OTPVerification.objects.filter(email=email, is_used=False).update(is_used=True)

            # Save new OTP
            OTPVerification.objects.create(email=email, otp_code=otp)

            # Send OTP via email
            try:
                send_mail(
                    subject='ScanPDF - Password Reset OTP',
                    message=f'Your OTP for password reset is: {otp}\n\nThis OTP is valid for 10 minutes.\n\nIf you did not request this, please ignore this email.',
                    from_email=settings.DEFAULT_FROM_EMAIL,
                    recipient_list=[email],
                    fail_silently=False,
                    html_message=f"""
                    <div style="font-family: 'Segoe UI', Arial, sans-serif; max-width: 480px; margin: 0 auto; padding: 40px 30px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); border-radius: 20px;">
                        <div style="background: white; border-radius: 16px; padding: 40px 30px; text-align: center; box-shadow: 0 20px 60px rgba(0,0,0,0.15);">
                            <h1 style="font-size: 24px; font-weight: 900; color: #1e293b; margin-bottom: 8px;">🔐 Password Reset</h1>
                            <p style="color: #64748b; font-size: 14px; margin-bottom: 30px;">Use this OTP to reset your ScanPDF account password</p>
                            <div style="background: linear-gradient(135deg, #4f46e5, #7c3aed); border-radius: 16px; padding: 24px; margin-bottom: 24px;">
                                <span style="font-size: 36px; font-weight: 900; letter-spacing: 12px; color: white; font-family: monospace;">{otp}</span>
                            </div>
                            <p style="color: #94a3b8; font-size: 12px; margin-bottom: 4px;">This OTP is valid for <strong>10 minutes</strong>.</p>
                            <p style="color: #cbd5e1; font-size: 11px;">If you didn't request this, please ignore this email.</p>
                        </div>
                        <p style="text-align: center; color: rgba(255,255,255,0.7); font-size: 11px; margin-top: 20px;">ScanPDF — All-in-One PDF Tools</p>
                    </div>
                    """,
                )
                # Store email in session for the verify step
                request.session['otp_email'] = email
                return redirect('dynamic_qr:verify_otp')
            except Exception as e:
                error = f"Failed to send OTP email. Please check email configuration. ({str(e)})"

    return render(request, 'dynamic_qr/forgot_password.html', {
        'error': error,
        'message': message,
    })


# ═══════════════════════════════════════════════════════════════
# AUTH: VERIFY OTP
# ═══════════════════════════════════════════════════════════════
def dqr_verify_otp_view(request):
    """Verify the OTP sent to email."""
    email = request.session.get('otp_email')
    if not email:
        return redirect('dynamic_qr:forgot_password')

    error = None
    if request.method == 'POST':
        entered_otp = request.POST.get('otp', '').strip()

        # Find the latest unused OTP for this email
        otp_record = OTPVerification.objects.filter(
            email=email,
            is_used=False,
        ).order_by('-created_at').first()

        if not otp_record:
            error = "No valid OTP found. Please request a new one."
        elif otp_record.attempts >= 5:
            otp_record.is_used = True
            otp_record.save()
            error = "Too many failed attempts. Please request a new OTP."
        elif (timezone.now() - otp_record.created_at) > timedelta(minutes=10):
            otp_record.is_used = True
            otp_record.save()
            error = "OTP has expired. Please request a new one."
        elif otp_record.otp_code != entered_otp:
            otp_record.attempts += 1
            otp_record.save()
            error = f"Invalid OTP. {5 - otp_record.attempts} attempts remaining."
        else:
            # OTP verified
            otp_record.is_used = True
            otp_record.save()
            request.session['otp_verified'] = True
            return redirect('dynamic_qr:reset_password')

    return render(request, 'dynamic_qr/verify_otp.html', {
        'email': email,
        'error': error,
    })


# ═══════════════════════════════════════════════════════════════
# AUTH: RESET PASSWORD (after OTP verification)
# ═══════════════════════════════════════════════════════════════
def dqr_reset_password_view(request):
    """Reset password after successful OTP verification."""
    email = request.session.get('otp_email')
    verified = request.session.get('otp_verified')

    if not email or not verified:
        return redirect('dynamic_qr:forgot_password')

    error = None
    if request.method == 'POST':
        form = ResetPasswordForm(request.POST)
        if form.is_valid():
            try:
                user = User.objects.get(email=email)
                user.set_password(form.cleaned_data['new_password'])
                user.save()

                # Clean up session
                del request.session['otp_email']
                del request.session['otp_verified']

                return render(request, 'dynamic_qr/password_reset_success.html')
            except User.DoesNotExist:
                error = "User account not found."
        else:
            error = form.errors.as_text()
    else:
        form = ResetPasswordForm()

    return render(request, 'dynamic_qr/reset_password.html', {
        'form': form,
        'error': error,
    })


# ═══════════════════════════════════════════════════════════════
# DASHBOARD: Overview & Recent
# ═══════════════════════════════════════════════════════════════
@dqr_login_required
def dqr_dashboard_view(request):
    """Overview dashboard with stats and top 6 recent QRs."""
    try:
        all_qrs = DynamicQRCode.objects.filter(user=request.user).order_by('-created_at')
        recent_qrs = all_qrs[:6]
        
        total_active = all_qrs.filter(is_active=True).count()
        total_deactivated = all_qrs.filter(is_active=False).count()
        
        from django.db.models import Sum
        total_scans = all_qrs.aggregate(Sum('scan_count'))['scan_count__sum'] or 0

        return render(request, 'dynamic_qr/dashboard.html', {
            'qr_codes': recent_qrs,
            'total_active': total_active,
            'total_deactivated': total_deactivated,
            'total_scans': total_scans,
            'has_more': all_qrs.count() > 6
        })
    except Exception as e:
        if 'no such column' in str(e).lower():
            return redirect('dynamic_qr:repair_db')
        raise e


@dqr_login_required
def dqr_all_qrs_view(request):
    """Full list of all QR codes with pagination."""
    try:
        from django.core.paginator import Paginator
        all_qrs = DynamicQRCode.objects.filter(user=request.user).order_by('-created_at')
        
        paginator = Paginator(all_qrs, 12) # 12 per page
        page_number = request.GET.get('page')
        page_obj = paginator.get_page(page_number)
        
        return render(request, 'dynamic_qr/all_qrs.html', {
            'page_obj': page_obj,
            'total_count': all_qrs.count()
        })
    except Exception as e:
        if 'no such column' in str(e).lower():
            return redirect('dynamic_qr:repair_db')
        raise e


# ═══════════════════════════════════════════════════════════════
# DASHBOARD: Create dynamic QR code
# ═══════════════════════════════════════════════════════════════
@dqr_login_required
def dqr_create_view(request):
    """Full page to create a new dynamic QR code or process the creation via AJAX."""
    if request.method == 'GET':
        return render(request, 'dynamic_qr/create_qr.html')

    # POST handles the creation (AJAX)
    qr_name = request.POST.get('qr_name', '').strip()
    qr_type = request.POST.get('qr_type', 'url').strip()
    destination_url = request.POST.get('destination_url', '').strip()
    qr_data_json = request.POST.get('qr_data', '{}')
    fg_color = request.POST.get('fg_color', '#000000')
    bg_color = request.POST.get('bg_color', '#ffffff')
    body_style = request.POST.get('body_style', 'square')
    eye_style = request.POST.get('eye_style', 'square')
    ball_style = request.POST.get('ball_style', 'square')
    logo = request.FILES.get('logo')
    file_content = request.FILES.get('file_content')

    if not qr_name:
        return JsonResponse({'error': 'Please enter a name for your QR code.'}, status=400)
    
    try:
        qr_data = json.loads(qr_data_json)
    except:
        qr_data = {}

    qr = DynamicQRCode.objects.create(
        user=request.user,
        qr_name=qr_name,
        qr_type=qr_type,
        destination_url=destination_url,
        qr_data=qr_data,
        fg_color=fg_color,
        bg_color=bg_color,
        body_style=body_style,
        eye_style=eye_style,
        ball_style=ball_style,
        logo=logo,
        file_content=file_content,
    )

    # Build the redirect URL that the QR code will point to
    redirect_url = request.build_absolute_uri(f'/qr/r/{qr.short_code}/')

    return JsonResponse({
        'success': True,
        'qr_id': str(qr.id),
        'short_code': qr.short_code,
        'redirect_url': redirect_url,
        'qr_name': qr.qr_name,
        'qr_type': qr.qr_type,
    })

    return JsonResponse({'error': 'POST required.'}, status=405)


# ═══════════════════════════════════════════════════════════════
# DASHBOARD: Edit/Update dynamic QR code
# ═══════════════════════════════════════════════════════════════
@dqr_login_required
def dqr_edit_view(request, qr_id):
    """Edit page for a dynamic QR code."""
    qr = get_object_or_404(DynamicQRCode, id=qr_id, user=request.user)

    if request.method == 'POST':
        qr_name = request.POST.get('qr_name', '').strip()
        qr_type = request.POST.get('qr_type', qr.qr_type).strip()
        destination_url = request.POST.get('destination_url', '').strip()
        qr_data_json = request.POST.get('qr_data', '{}')

        fg_color = request.POST.get('fg_color', qr.fg_color)
        bg_color = request.POST.get('bg_color', qr.bg_color)
        body_style = request.POST.get('body_style', qr.body_style)
        eye_style = request.POST.get('eye_style', qr.eye_style)
        ball_style = request.POST.get('ball_style', qr.ball_style)
        is_active = request.POST.get('is_active', 'true') == 'true'
        logo = request.FILES.get('logo')
        file_content = request.FILES.get('file_content')

        if not qr_name:
            return JsonResponse({'error': 'QR name is required.'}, status=400)

        qr.qr_name = qr_name
        qr.qr_type = qr_type
        qr.destination_url = destination_url
        try:
            qr.qr_data = json.loads(qr_data_json)
        except:
            pass
            
        qr.fg_color = fg_color
        qr.bg_color = bg_color
        qr.body_style = body_style
        qr.eye_style = eye_style
        qr.ball_style = ball_style
        qr.is_active = is_active
        if logo:
            qr.logo = logo
        if file_content:
            qr.file_content = file_content
        qr.save()

        # If AJAX request, return JSON
        if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
            redirect_url = request.build_absolute_uri(f'/qr/r/{qr.short_code}/')
            return JsonResponse({
                'success': True,
                'qr_id': str(qr.id),
                'short_code': qr.short_code,
                'redirect_url': redirect_url,
                'qr_name': qr.qr_name,
                'qr_type': qr.qr_type,
            })

        return redirect('dynamic_qr:dashboard')

    redirect_url = request.build_absolute_uri(f'/qr/r/{qr.short_code}/')
    return render(request, 'dynamic_qr/edit_qr.html', {
        'qr': qr,
        'redirect_url': redirect_url,
    })


@dqr_login_required
@require_POST
def dqr_delete_view(request, qr_id):
    """Delete a dynamic QR code."""
    qr = get_object_or_404(DynamicQRCode, id=qr_id, user=request.user)
    qr.delete()
    if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
        return JsonResponse({'success': True})
    return redirect('dynamic_qr:dashboard')

@dqr_login_required
@require_POST
def dqr_toggle_status(request, qr_id):
    """Toggle the Active/Inactive status of a QR code."""
    qr = get_object_or_404(DynamicQRCode, id=qr_id, user=request.user)
    qr.is_active = not qr.is_active
    qr.save()
    return JsonResponse({'success': True, 'is_active': qr.is_active})


@dqr_login_required
def dqr_analytics_view(request, qr_id):
    """Detailed analytics for a specific QR code with advanced filtering and chart data."""
    qr = get_object_or_404(DynamicQRCode, id=qr_id, user=request.user)
    
    from django.db.models import Count
    from django.db.models.functions import TruncDate
    from django.db import connection, OperationalError
    
    from django.core.paginator import Paginator
    
    selected_range = request.GET.get('range', '7days')
    now = timezone.now()
    
    if selected_range == 'today':
        start_date = now.replace(hour=0, minute=0, second=0, microsecond=0)
    elif selected_range == '7days':
        start_date = now - timedelta(days=7)
    elif selected_range == '1month' or selected_range == '30days':
        start_date = now - timedelta(days=30)
    elif selected_range == '28days':
        start_date = now - timedelta(days=28)
    elif selected_range == '6months':
        start_date = now - timedelta(days=180)
    elif selected_range == '12months':
        start_date = now - timedelta(days=365)
    else:
        start_date = now - timedelta(days=7)
        
    def get_data():
        base_query = qr.analytics.filter(timestamp__gte=start_date)
        daily_scans = list(base_query.annotate(date=TruncDate('timestamp')).values('date').annotate(count=Count('id')).order_by('date'))
        browser_stats = list(base_query.values('browser').annotate(count=Count('id')).order_by('-count')[:5])
        device_stats = list(base_query.values('device_type').annotate(count=Count('id')).order_by('-count'))
        os_stats = list(base_query.values('os').annotate(count=Count('id')).order_by('-count')[:5])
        # For pagination, we need the queryset, not a slice
        recent_scans_qs = base_query.order_by('-timestamp')
        return daily_scans, browser_stats, device_stats, os_stats, recent_scans_qs

    try:
        daily_scans, browser_stats, device_stats, os_stats, recent_scans_qs = get_data()
    except OperationalError:
        daily_scans, browser_stats, device_stats, os_stats, recent_scans_qs = [], [], [], [], []

    # Pagination Logic
    paginator = Paginator(recent_scans_qs, 10)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)

    chart_labels = [d['date'].strftime('%b %d') for d in daily_scans]
    chart_data = [d['count'] for d in daily_scans]
    device_labels = [d['device_type'] if d['device_type'] else 'Unknown' for d in device_stats]
    device_data = [d['count'] for d in device_stats]
    browser_labels = [d['browser'] if d['browser'] else 'Other' for d in browser_stats]
    browser_data = [d['count'] for d in browser_stats]

    return render(request, 'dynamic_qr/analytics.html', {
        'qr': qr,
        'selected_range': selected_range,
        'total_range_scans': sum(chart_data),
        'js_labels': json.dumps(chart_labels),
        'js_data': json.dumps(chart_data),
        'js_device_labels': json.dumps(device_labels),
        'js_device_data': json.dumps(device_data),
        'js_browser_labels': json.dumps(browser_labels),
        'js_browser_data': json.dumps(browser_data),
        'page_obj': page_obj,
        'browser_stats': browser_stats,
        'device_stats': device_stats,
        'os_stats': os_stats,
    })


@dqr_login_required
def dqr_details_view(request, qr_id):
    """Quick details view with download options."""
    qr = get_object_or_404(DynamicQRCode, id=qr_id, user=request.user)
    redirect_url = request.build_absolute_uri(f'/qr/r/{qr.short_code}/')
    return render(request, 'dynamic_qr/details.html', {
        'qr': qr,
        'redirect_url': redirect_url,
    })


# ═══════════════════════════════════════════════════════════════
# REDIRECT: Short code → Destination URL
# ═══════════════════════════════════════════════════════════════
def dqr_redirect_view(request, short_code):
    """
    When someone scans the dynamic QR code, they hit this URL.
    De-duplicates hits to prevent double-counting from pre-fetchers.
    """
    qr = get_object_or_404(DynamicQRCode, short_code=short_code)

    if not qr.is_active:
        return render(request, 'dynamic_qr/qr_disabled.html', {'qr': qr})

    # Logic for Logging (only once every 5 seconds per session)
    now_ts = timezone.now().timestamp()
    last_ts = request.session.get(f'qr_last_hit_{qr.id}', 0)
    
    if (now_ts - last_ts) >= 5:
        # 1. Increment Scan Count
        qr.increment_scan()
        request.session[f'qr_last_hit_{qr.id}'] = now_ts
        
        # 2. Log Detailed Analytics
        ua = request.META.get('HTTP_USER_AGENT', '').lower()
        ip = request.META.get('REMOTE_ADDR')
        
        # Simple Manual Parsing
        browser = 'Other'
        if 'chrome' in ua: browser = 'Chrome'
        elif 'safari' in ua: browser = 'Safari'
        elif 'firefox' in ua: browser = 'Firefox'
        elif 'edge' in ua: browser = 'Edge'
        
        os = 'Unknown'
        if 'windows' in ua: os = 'Windows'
        elif 'android' in ua: os = 'Android'
        elif 'iphone' in ua or 'ipad' in ua: os = 'iOS'
        elif 'mac' in ua: os = 'macOS'
        elif 'linux' in ua: os = 'Linux'
        
        device = 'Desktop'
        if 'mobile' in ua or 'android' in ua or 'iphone' in ua: device = 'Mobile'
        elif 'tablet' in ua or 'ipad' in ua: device = 'Tablet'

        from .models import QRAnalytics
        from django.db import connection, OperationalError
        try:
            QRAnalytics.objects.create(qr_code=qr, ip_address=ip, user_agent=ua[:500], browser=browser, os=os, device_type=device)
        except OperationalError:
            with connection.cursor() as cursor:
                cursor.execute('CREATE TABLE IF NOT EXISTS "dynamic_qr_qranalytics" ("id" integer NOT NULL PRIMARY KEY AUTOINCREMENT, "timestamp" datetime NOT NULL, "ip_address" char(39) NULL, "user_agent" text NULL, "browser" varchar(50) NULL, "os" varchar(50) NULL, "device_type" varchar(50) NULL, "country" varchar(100) NOT NULL DEFAULT "Unknown", "city" varchar(100) NOT NULL DEFAULT "Unknown", "qr_code_id" uuid NOT NULL REFERENCES "dynamic_qr_dynamicqrcode" ("id") DEFERRABLE INITIALLY DEFERRED);')
            try: QRAnalytics.objects.create(qr_code=qr, ip_address=ip, user_agent=ua[:500], browser=browser, os=os, device_type=device)
            except: pass
        except: pass

    # Determine redirect behavior
    redirect_types = [
        'url', 'whatsapp', 'youtube', 'facebook', 'instagram', 'telegram', 
        'tiktok', 'x-twitter', 'snapchat', 'pinterest', 'linkedin',
        'pdf', 'audio', 'video', 'image', 'pptx', 'excel', 'word',
        'google-review', 'google-forms', 'google-doc', 'google-sheets',
        'play-market', 'app-store', 'paypal', 'etsy', 'amazon', 'venmo', 
        'upi', 'crypto', 'spotify', 'link-list', 'custom-url', 'office-365'
    ]

    target_url = None
    if qr.file_content:
        target_url = qr.file_content.url
    elif qr.destination_url:
        target_url = qr.destination_url

    if qr.qr_type in redirect_types and target_url:
        return HttpResponseRedirect(target_url)
    
    # For other types (text, wifi, vcard, calendar, etc.), show the content on a clean landing page
    return render(request, 'dynamic_qr/landing.html', {
        'qr': qr,
        'is_preview': False
    })


# ═══════════════════════════════════════════════════════════════
# API: Check dynamic QR auth status (used by QR generator page)
# ═══════════════════════════════════════════════════════════════
def dqr_auth_status(request):
    """Return whether the user is logged in for dynamic QR features."""
    return JsonResponse({
        'authenticated': request.user.is_authenticated,
        'username': request.user.username if request.user.is_authenticated else None,
    })


# ═══════════════════════════════════════════════════════════════
# API: Generate dynamic QR image (reuses existing QR engine)
# ═══════════════════════════════════════════════════════════════
@dqr_login_required
def dqr_generate_image(request):
    """Generate the QR code image for a dynamic QR entry."""
    data = request.POST if request.method == 'POST' else request.GET

    from converter.utils import generate_qr_code, get_output_path
    from converter.views import create_cleanup_response

    text = data.get('text', '')
    fg_color = data.get('fg_color', '#000000')
    bg_color = data.get('bg_color', '#ffffff')
    style = data.get('style', 'square')
    eye_style = data.get('eye_style', 'square')
    ball_style = data.get('ball_style', 'square')
    output_format = data.get('output_format', 'png')

    if not text:
        return JsonResponse({'error': 'No QR content provided.'}, status=400)

    try:
        from converter.utils import save_uploaded_file
        logo_path = None
        if 'logo' in request.FILES:
            logo_path = save_uploaded_file(request.FILES['logo'])

        output_path = generate_qr_code(
            text,
            fg_color=fg_color,
            bg_color=bg_color,
            style=style,
            eye_style=eye_style,
            ball_style=ball_style,
            logo_path=logo_path,
            output_format=output_format,
        )

        if logo_path and os.path.exists(logo_path):
            try:
                os.remove(logo_path)
            except OSError:
                pass

        ct = 'image/png'
        if output_format.lower() in ('jpg', 'jpeg'):
            ct = 'image/jpeg'
        elif output_format.lower() == 'svg':
            ct = 'image/svg+xml'

        return create_cleanup_response(output_path, content_type=ct)
    except Exception as e:
        return JsonResponse({'error': str(e)}, status=500)


@dqr_login_required
def dqr_download_view(request, qr_id):
    """Generate and return the QR image for a specific dynamic QR code in requested format."""
    qr = get_object_or_404(DynamicQRCode, id=qr_id, user=request.user)
    fmt = request.GET.get('format', 'png').lower()
    if fmt not in ('png', 'jpg', 'jpeg', 'svg'):
        fmt = 'png'
        
    from converter.utils import generate_qr_code
    from converter.views import create_cleanup_response

    # Redefine redirect URL pointing to our redirector
    redirect_url = request.build_absolute_uri(f'/qr/r/{qr.short_code}/')
    
    # Path to existing logo if any
    logo_path = None
    if qr.logo and os.path.exists(qr.logo.path):
        logo_path = qr.logo.path

    # Use the helper from converter.utils
    output_path = generate_qr_code(
        redirect_url,
        fg_color=qr.fg_color,
        bg_color=qr.bg_color,
        style=qr.body_style,
        eye_style=qr.eye_style,
        ball_style=qr.ball_style,
        logo_path=logo_path,
        output_format=fmt
    )
    
    ct = 'image/png'
    if fmt in ('jpg', 'jpeg'): ct = 'image/jpeg'
    elif fmt == 'svg': ct = 'image/svg+xml'
    
    response = create_cleanup_response(output_path, content_type=ct)
    response['Content-Disposition'] = f'attachment; filename="QR_{qr.short_code}.{fmt}"'
    return response
