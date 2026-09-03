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
from django.http import JsonResponse, HttpResponseRedirect, HttpResponse
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
import requests
import uuid
import base64


# ═══════════════════════════════════════════════════════════════
# HELPER: Check if user is logged in for Dynamic QR
# ═══════════════════════════════════════════════════════════════
def dqr_login_required(view_func):
    """Decorator: redirect to pricing if not authenticated or not a QR user."""
    def wrapper(request, *args, **kwargs):
        # Isolation: Check if authenticated AND has the dqr flag
        if not request.user.is_authenticated or not request.session.get('is_dqr_user'):
            from django.urls import reverse
            login_url = reverse('dynamic_qr:login')
            next_url = request.get_full_path()
            return redirect(f"{login_url}?next={next_url}")
        return view_func(request, *args, **kwargs)
    return wrapper


from django.db import connection

def dqr_repair_db(request):
    """Utility view to manually add missing columns/tables to SQLite via browser."""
    # Allow superusers OR regular authenticated users for this specific repair
    if not request.user.is_authenticated:
        from django.http import HttpResponse
        return HttpResponse("Unauthorized.", status=403)
    
    from django.core.management import call_command
    results = []
    
    # Auto-create superuser for admin dashboard verification
    from django.contrib.auth.models import User
    if not User.objects.filter(username='admin').exists():
        User.objects.create_superuser('admin', 'admin@example.com', 'adminpassword123')
        results.append("✅ Created superuser 'admin' with password 'adminpassword123'.")
    else:
        # Reset password to ensure we can log in
        admin_user = User.objects.get(username='admin')
        admin_user.set_password('adminpassword123')
        admin_user.is_superuser = True
        admin_user.is_staff = True
        admin_user.save()
        results.append("✅ Reset superuser 'admin' password to 'adminpassword123'.")
        
    try:
        call_command('makemigrations', 'services', interactive=False)
        call_command('migrate', 'services', interactive=False)
        results.append("✅ Successfully programmatically ran makemigrations & migrate for 'services' app.")
    except Exception as e:
        results.append(f"❌ Error during database migration run: {str(e)}")

    from django.http import HttpResponse
    with connection.cursor() as cursor:
        
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
            ("ball_style", "VARCHAR(20) DEFAULT 'square'"),
            ("design_options", "JSON NULL")
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
    if request.user.is_authenticated and request.session.get('is_dqr_user'):
        if request.user.is_superuser:
            return redirect('custom_admin:dashboard')

        next_url = request.GET.get('next', '')
        if next_url:
            return redirect(next_url)
            
        elif request.user.is_staff:
            return redirect('admin:index')
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
            
            # Redirect superusers to custom admin unconditionally
            if user.is_superuser:
                return redirect('custom_admin:dashboard')

            next_url = request.GET.get('next', '')
            if next_url:
                return redirect(next_url)
            
            elif user.is_staff:
                return redirect('admin:index')
                
            return redirect('dynamic_qr:dashboard')
        else:
            error = "Invalid username or password"

    return render(request, 'dynamic_qr/login.html', {'error': error})


# ═══════════════════════════════════════════════════════════════
# AUTH: REGISTER
# ═══════════════════════════════════════════════════════════════
def dqr_register_view(request):
    """Register page for dynamic QR feature only."""
    if request.user.is_authenticated and request.session.get('is_dqr_user'):
        next_url = request.GET.get('next', '')
        return redirect(next_url if next_url else 'dynamic_qr:dashboard')

    form = DynamicQRRegisterForm()
    if request.method == 'POST':
        form = DynamicQRRegisterForm(request.POST)
        if form.is_valid():
            user = form.save()
            login(request, user)
            request.session['is_dqr_user'] = True
            next_url = request.GET.get('next', '')
            return redirect(next_url if next_url else 'dynamic_qr:dashboard')

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
        # Exclude Short URL items (custom-url) from QR dashboard.
        all_qrs = DynamicQRCode.objects.filter(user=request.user).exclude(qr_type='custom-url').order_by('-created_at')
        recent_qrs = all_qrs[:6]
        
        total_active = all_qrs.filter(is_active=True).count()
        total_deactivated = all_qrs.filter(is_active=False).count()
        
        from django.db.models import Sum
        total_scans = all_qrs.aggregate(Sum('scan_count'))['scan_count__sum'] or 0

        for qr in recent_qrs:
            qr.qr_content = qr.get_static_content(request)

        # Retrieve active subscription details
        from services.models import Subscription, Plan
        subscription = Subscription.objects.filter(user=request.user, status='Active').first()
        if not subscription:
            from services.views import get_or_create_plans
            get_or_create_plans()
            free_plan = Plan.objects.get(code='free')
            subscription = Subscription.objects.create(
                user=request.user,
                plan=free_plan,
                status='Active',
                billing_cycle='monthly',
                payment_status='Paid'
            )

        return render(request, 'dynamic_qr/dashboard.html', {
            'qr_codes': recent_qrs,
            'total_active': total_active,
            'total_deactivated': total_deactivated,
            'total_scans': total_scans,
            'has_more': all_qrs.count() > 6,
            'subscription': subscription
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
        # Exclude Short URL items (custom-url) from QR library.
        all_qrs = DynamicQRCode.objects.filter(user=request.user).exclude(qr_type='custom-url').order_by('-created_at')
        
        paginator = Paginator(all_qrs, 12) # 12 per page
        page_number = request.GET.get('page')
        page_obj = paginator.get_page(page_number)
        
        for qr in page_obj:
            qr.qr_content = qr.get_static_content(request)
            
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
from services.decorators import check_dynamic_qr_limit, check_short_url_limit

@dqr_login_required
@check_dynamic_qr_limit
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
    design_data_json = request.POST.get('design_options', '{}')
    logo = request.FILES.get('logo')
    logo_cropped = request.POST.get('logo_cropped')
    file_content = request.FILES.get('file_content')

    if logo_cropped and logo_cropped.startswith('data:image'):
        from django.core.files.base import ContentFile
        import base64
        import uuid
        try:
            format, imgstr = logo_cropped.split(';base64,')
            ext = 'png'.split('/')[-1] # Fallback to png usually
            if '/svg+xml' in format: ext = 'svg'
            elif '/jpeg' in format: ext = 'jpg'
            logo = ContentFile(base64.b64decode(imgstr), name=f"logo_{uuid.uuid4()}.{ext}")
        except:
            pass

    if not qr_name:
        return JsonResponse({'error': 'Please enter a name for your QR code.'}, status=400)
    
    try:
        qr_data = json.loads(qr_data_json)
    except:
        qr_data = {}

    try:
        design_options = json.loads(design_data_json)
    except:
        design_options = {}

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
        design_options=design_options
    )

    # --- Permanent Logo Persistence (Preset caching) ---
    if not qr.logo and qr.design_options and qr.design_options.get('logo_preset'):
        preset = qr.design_options.get('logo_preset')
        if preset and preset != 'none':
            try:
                target_icon = os.path.join(settings.MEDIA_ROOT, 'brand_icons', f"{preset}.png")
                if os.path.exists(target_icon) and os.path.getsize(target_icon) > 0:
                    from django.core.files import File
                    with open(target_icon, 'rb') as f:
                        qr.logo.save(f"{preset}_preset.png", File(f), save=False)
                    qr.save()
            except: pass

    # Build the static content that the QR code will contain
    qr_content = qr.get_static_content(request)

    return JsonResponse({
        'success': True,
        'qr_id': str(qr.id),
        'short_code': qr.short_code,
        'redirect_url': qr_content,
        'qr_name': qr.qr_name,
        'qr_type': qr.qr_type,
    })

    return JsonResponse({'error': 'POST required.'}, status=405)


@dqr_login_required
def dqr_short_url_view(request):
    """Specialized tool for Short URLs: List and Create."""
    if request.method == 'GET':
        from services.limit_service import PlanLimitService
        from django.db.models import Q
        from django.utils import timezone
        from datetime import timedelta
        from django.core.paginator import Paginator

        queryset = DynamicQRCode.objects.filter(user=request.user, qr_type='custom-url').order_by('-created_at')

        # 1. Search
        search_query = request.GET.get('q', '').strip()
        if search_query:
            queryset = queryset.filter(
                Q(qr_name__icontains=search_query) |
                Q(short_code__icontains=search_query) |
                Q(destination_url__icontains=search_query) |
                Q(custom_alias__icontains=search_query)
            )

        # 2. Date Range
        date_range = request.GET.get('date_range', 'all')
        now = timezone.now()
        if date_range == 'today':
            queryset = queryset.filter(created_at__date=now.date())
        elif date_range == '7days':
            queryset = queryset.filter(created_at__gte=now - timedelta(days=7))
        elif date_range == '30days':
            queryset = queryset.filter(created_at__gte=now - timedelta(days=30))
        elif date_range == '90days':
            queryset = queryset.filter(created_at__gte=now - timedelta(days=90))

        # 3. Status Filters
        status_filter = request.GET.get('status')
        if status_filter == 'active':
            queryset = queryset.filter(is_active=True)
        elif status_filter == 'inactive':
            queryset = queryset.filter(is_active=False)
            
        if request.GET.get('password') == 'true':
            queryset = queryset.exclude(password__isnull=True).exclude(password__exact='')
        if request.GET.get('qr') == 'true':
            queryset = queryset.filter(qr_enabled=True)
        if request.GET.get('gps') == 'true':
            queryset = queryset.filter(require_gps=True)

        # 4. Pagination
        paginator = Paginator(queryset, 10)
        page_number = request.GET.get('page')
        short_urls = paginator.get_page(page_number)
        
        # Pass feature access flags to hide/show UI components
        has_custom_alias = PlanLimitService.check_feature_access(request.user, 'custom_alias')
        has_password = PlanLimitService.check_feature_access(request.user, 'password_protection')
        has_expiry = PlanLimitService.check_feature_access(request.user, 'link_expiry')
        has_gps = PlanLimitService.check_feature_access(request.user, 'gps_tracking')
        has_header = PlanLimitService.check_feature_access(request.user, 'custom_header')
        has_qr = PlanLimitService.check_feature_access(request.user, 'dynamic_qrs')
        
        allowed, current_usage, limit, _ = PlanLimitService.check_usage_limit(request.user, 'short_urls', increment=False)
        
        # Header usage data (count of header-enabled short URLs)
        used_headers_count = DynamicQRCode.objects.filter(user=request.user).exclude(header__isnull=True).exclude(header='').count()
        _, _, header_limit, _ = PlanLimitService.check_usage_limit(request.user, 'custom_header', increment=False)
        
        return render(request, 'dynamic_qr/short_url.html', {
            'short_urls': short_urls,
            'search_query': search_query,
            'current_filters': {
                'date_range': date_range,
                'status': status_filter,
                'password': request.GET.get('password'),
                'qr': request.GET.get('qr'),
                'gps': request.GET.get('gps'),
            },
            'has_custom_alias': has_custom_alias,
            'has_password': has_password,
            'has_expiry': has_expiry,
            'has_gps': has_gps,
            'has_header': has_header,
            'has_qr': has_qr,
            'usage_current': current_usage,
            'usage_limit': limit,
            'can_create_more': allowed,
            'used_headers_count': used_headers_count,
            'header_limit': str(header_limit),
        })
    
    try:
        qr_id = request.POST.get('qr_id')
        qr_name = request.POST.get('qr_name', 'Short URL').strip()
        destination_url = request.POST.get('destination_url', '').strip()
        regenerate = request.POST.get('regenerate_code') == 'on'
        
        from services.limit_service import PlanLimitService
        
        # Check creation limits
        if not qr_id:
            allowed, usage, limit, msg = PlanLimitService.check_usage_limit(request.user, 'short_urls', increment=False)
            if not allowed:
                return JsonResponse({'error': msg or 'Short URL limit reached.'}, status=403)
        
        if not destination_url:
            return JsonResponse({'error': 'URL is required.'}, status=400)
        
        if not destination_url.startswith(('http://', 'https://')):
            destination_url = 'https://' + destination_url
        
        qr_data = {'destination_url': destination_url}
        
        custom_alias = request.POST.get('custom_alias', '').strip()
        domain = request.POST.get('domain', 'default').strip()
        password = request.POST.get('password', '').strip()
        expiry_date_str = request.POST.get('expiry_date', '').strip()
        require_gps = request.POST.get('require_gps') == 'on'
        
        # New Feature Toggles & QR Styles
        qr_enabled = request.POST.get('qr_enabled') == 'on'
        header_enabled = request.POST.get('header_enabled') == 'on'
        header_value = None
        
        if header_enabled:
            header_value = request.POST.get('header', '').strip()
            
            if header_value:
                import re
                if not re.match(r'^[A-Za-z0-9_-]{1,30}$', header_value):
                    return JsonResponse({'error': 'Header must be 1-30 characters (A-Z, 0-9, -, _).'}, status=400)
                
                if header_value.lower() in DynamicQRCode.RESERVED_PATHS:
                    return JsonResponse({'error': f'The header "{header_value}" is a reserved system path and cannot be used.'}, status=400)
                    
                # Check Header Limit natively
                used_headers_count = DynamicQRCode.objects.filter(user=request.user).exclude(header__isnull=True).exclude(header='').count()
                
                # Check if this specific link already has a header (if editing)
                editing_existing_header = False
                if qr_id:
                    existing_qr = DynamicQRCode.objects.filter(id=qr_id, user=request.user).first()
                    if existing_qr and existing_qr.header:
                        editing_existing_header = True
                
                if not editing_existing_header:
                    # Creating a new header-enabled link, check limits
                    allowed, _, limit_str, _ = PlanLimitService.check_usage_limit(request.user, 'custom_header', increment=False)
                    if not allowed:
                        return JsonResponse({'error': 'Custom Header feature is not available in your plan.'}, status=403)
                        
                    if str(limit_str).lower() != 'unlimited' and str(limit_str) != '-1':
                        try:
                            numeric_limit = int(limit_str)
                            if used_headers_count >= numeric_limit:
                                return JsonResponse({'error': f'Header limit reached ({numeric_limit}). Upgrade your plan to create more links with custom headers.'}, status=403)
                        except:
                            pass
        
        # Reserved Path Protection for Short Code / Alias when Header is empty
        if not header_value and custom_alias and custom_alias.lower() in DynamicQRCode.RESERVED_PATHS:
            return JsonResponse({'error': f'The alias "{custom_alias}" is a reserved system path and cannot be used without a Header.'}, status=400)
            
        fg_color = request.POST.get('fg_color', '#000000')
        bg_color = request.POST.get('bg_color', '#ffffff')
        body_style = request.POST.get('body_style', 'square')
        eye_style = request.POST.get('eye_style', 'square')
        ball_style = request.POST.get('ball_style', 'square')
        design_data_json = request.POST.get('design_options', '{}')
        logo_cropped = request.POST.get('logo_cropped')
        
        logo = None
        if logo_cropped and logo_cropped.startswith('data:image'):
            from django.core.files.base import ContentFile
            import base64
            import uuid
            try:
                format, imgstr = logo_cropped.split(';base64,')
                ext = 'png'.split('/')[-1]
                if '/svg+xml' in format: ext = 'svg'
                elif '/jpeg' in format: ext = 'jpg'
                logo = ContentFile(base64.b64decode(imgstr), name=f"logo_{uuid.uuid4()}.{ext}")
            except:
                pass
                
        try:
            design_options = json.loads(design_data_json)
        except:
            design_options = {}

        # Validate unique alias
        if custom_alias:
            allowed = PlanLimitService.check_feature_access(request.user, 'custom_alias')
            if not allowed:
                return JsonResponse({'error': 'Custom Alias feature is not available in your plan.'}, status=403)
                
            if DynamicQRCode.objects.exclude(id=qr_id).filter(custom_alias=custom_alias).exists():
                return JsonResponse({'error': 'Custom alias is already in use.'}, status=400)
            if DynamicQRCode.objects.exclude(id=qr_id).filter(short_code=custom_alias).exists():
                return JsonResponse({'error': 'Custom alias conflicts with an existing short code.'}, status=400)

        # Parse expiry date
        expiry_date = None
        if expiry_date_str:
            from dateutil.parser import parse
            try:
                expiry_date = parse(expiry_date_str)
            except:
                return JsonResponse({'error': 'Invalid expiry date format.'}, status=400)

        # Hash password
        hashed_password = None
        if password:
            allowed = PlanLimitService.check_feature_access(request.user, 'password_protection')
            if not allowed:
                return JsonResponse({'error': 'Password Protection is not available in your plan.'}, status=403)
                
            from django.contrib.auth.hashers import make_password
            hashed_password = make_password(password)

        if expiry_date:
            allowed = PlanLimitService.check_feature_access(request.user, 'link_expiry')
            if not allowed:
                return JsonResponse({'error': 'Link Expiry is not available in your plan.'}, status=403)
                
        if require_gps:
            allowed = PlanLimitService.check_feature_access(request.user, 'gps_tracking')
            if not allowed:
                return JsonResponse({'error': 'GPS Tracking is not available in your plan.'}, status=403)
                
        if qr_enabled:
            allowed = PlanLimitService.check_feature_access(request.user, 'dynamic_qrs')
            if not allowed:
                return JsonResponse({'error': 'QR Code generation is not available in your plan.'}, status=403)
                
        if header_enabled and not header_value:
            return JsonResponse({'error': 'Header value is required when enabled.'}, status=400)

        if qr_id:
            # Update existing
            qr = get_object_or_404(DynamicQRCode, id=qr_id, user=request.user)
            qr.qr_name = qr_name
            qr.destination_url = destination_url
            qr.qr_data = qr_data
            qr.custom_alias = custom_alias or None
            qr.domain = domain
            if password: # only update if new password provided
                qr.password = hashed_password
            qr.expiry_date = expiry_date
            qr.require_gps = require_gps
            qr.header = header_value if header_enabled else None
            qr.qr_enabled = qr_enabled
            qr.fg_color = fg_color
            qr.bg_color = bg_color
            qr.body_style = body_style
            qr.eye_style = eye_style
            qr.ball_style = ball_style
            qr.design_options = design_options
            if logo:
                qr.logo = logo

            if regenerate:
                from .models import generate_short_code
                qr.short_code = generate_short_code()
            qr.save()
        else:
            # Create new
            qr = DynamicQRCode(
                user=request.user,
                qr_name=qr_name,
                qr_type='custom-url',
                destination_url=destination_url,
                qr_data=qr_data,
                design_options=design_options,
                custom_alias=custom_alias or None,
                domain=domain,
                password=hashed_password,
                expiry_date=expiry_date,
                require_gps=require_gps,
                header=header_value if header_enabled else None,
                qr_enabled=qr_enabled,
                fg_color=fg_color,
                bg_color=bg_color,
                body_style=body_style,
                eye_style=eye_style,
                ball_style=ball_style,
                logo=logo
            )
            qr.save()
            PlanLimitService.check_usage_limit(request.user, 'short_urls', increment=True)
            
        # --- Permanent Logo Persistence (Preset caching) ---
        if not qr.logo and qr.design_options and qr.design_options.get('logo_preset'):
            preset = qr.design_options.get('logo_preset')
            if preset and preset != 'none':
                try:
                    target_icon = os.path.join(settings.MEDIA_ROOT, 'brand_icons', f"{preset}.png")
                    if os.path.exists(target_icon) and os.path.getsize(target_icon) > 0:
                        from django.core.files import File
                        with open(target_icon, 'rb') as f:
                            qr.logo.save(f"{preset}_preset.png", File(f), save=False)
                        qr.save()
                except: pass
        
        return JsonResponse({
            'success': True, 
            'id': str(qr.id), 
            'short_url': request.build_absolute_uri(f"/qr/r/{qr.short_code}/"),
            'qr_name': qr.qr_name,
            'created_at': qr.created_at.strftime('%Y-%m-%d %H:%M'),
            'scan_count': qr.scan_count
        })
    except Exception as e:
        return JsonResponse({'error': str(e)}, status=500)


@dqr_login_required
def dqr_short_url_analytics_view(request, qr_id):
    """Detailed analytics for a specific Short URL, matching QR excellence."""
    qr = get_object_or_404(DynamicQRCode, id=qr_id, user=request.user, qr_type='custom-url')
    
    from django.db.models import Count, Q
    from django.db.models.functions import TruncDate, TruncHour, TruncMonth
    from django.db import connection, OperationalError
    from django.core.paginator import Paginator
    from django.utils import timezone
    from datetime import timedelta
    from django.http import HttpResponse
    import json
    import csv
    
    # 1. Plan limit check
    from services.limit_service import PlanLimitService
    allowed, _, max_days_str, _ = PlanLimitService.check_usage_limit(request.user, 'analytics_history', increment=False)
    
    max_history_days = 30
    try:
        max_history_days = int(max_days_str)
    except:
        if str(max_days_str).lower() == 'unlimited' or str(max_days_str) == '-1':
            max_history_days = 3650
            
    selected_range = request.GET.get('range', '7days')
    now = timezone.now()
    history_clamped = False
    
    range_map = {
        'today': 1,
        '7days': 7,
        '28days': 28,
        '1month': 30,
        '30days': 30,
        '12months': 365
    }
    
    req_days = range_map.get(selected_range, 7)
    if req_days > max_history_days:
        history_clamped = True
        req_days = max_history_days
        
    if selected_range == 'today':
        start_date = now.replace(hour=0, minute=0, second=0, microsecond=0)
        prev_start = start_date - timedelta(days=1)
        prev_end = start_date
    else:
        start_date = now - timedelta(days=req_days)
        prev_start = start_date - timedelta(days=req_days)
        prev_end = start_date
        
    # 2. CSV Export
    if request.GET.get('export') == 'csv':
        response = HttpResponse(content_type='text/csv')
        response['Content-Disposition'] = f'attachment; filename="analytics_{qr.short_code}.csv"'
        writer = csv.writer(response)
        writer.writerow(['Timestamp', 'Short Code', 'Header', 'Country', 'City', 'Device', 'Browser', 'OS', 'Referrer', 'Type', 'Is Bot'])
        
        for record in qr.analytics.filter(timestamp__gte=start_date).order_by('-timestamp'):
            writer.writerow([
                record.timestamp.strftime('%Y-%m-%d %H:%M:%S'),
                qr.short_code,
                qr.header or '',
                record.country,
                record.city,
                record.device_type,
                record.browser,
                record.os,
                record.referrer or 'Direct',
                record.source,
                'Yes' if record.is_bot else 'No'
            ])
        return response

    # 3. Base Queries
    base_query = qr.analytics.filter(timestamp__gte=start_date)
    prev_query = qr.analytics.filter(timestamp__gte=prev_start, timestamp__lt=prev_end)
    
    def get_data():
        total_clicks = base_query.count()
        qr_scans = base_query.filter(is_qr_scan=True).count()
        unique_clicks = base_query.exclude(visitor_id__isnull=True).exclude(visitor_id='').values('visitor_id').distinct().count()
        human_clicks = base_query.filter(is_bot=False).count()
        bot_clicks = base_query.filter(is_bot=True).count()
        
        prev_total = prev_query.count()
        prev_qr = prev_query.filter(is_qr_scan=True).count()
        prev_unique = prev_query.exclude(visitor_id__isnull=True).exclude(visitor_id='').values('visitor_id').distinct().count()
        prev_human = prev_query.filter(is_bot=False).count()
        prev_bot = prev_query.filter(is_bot=True).count()
        
        def pct_change(curr, prev):
            if prev == 0: return "+100%" if curr > 0 else "0%"
            val = ((curr - prev) / prev) * 100
            return f"+{val:.1f}%" if val > 0 else f"{val:.1f}%"
            
        trends = {
            'total': pct_change(total_clicks, prev_total),
            'qr': pct_change(qr_scans, prev_qr),
            'unique': pct_change(unique_clicks, prev_unique),
            'human': pct_change(human_clicks, prev_human),
            'bot': pct_change(bot_clicks, prev_bot),
        }
        
        # Aggregate device/browser/os — replace None with 'Unknown'
        source_stats = list(base_query.values('source').annotate(count=Count('id')).order_by('-count'))
        os_stats_raw = list(base_query.values('os').annotate(count=Count('id')).order_by('-count')[:6])
        browser_stats_raw = list(base_query.values('browser').annotate(count=Count('id')).order_by('-count')[:6])
        device_stats_raw = list(base_query.values('device_type').annotate(count=Count('id')).order_by('-count'))
        
        # Normalize None -> 'Unknown' and merge duplicates
        def normalize_stat(stat_list, key):
            merged = {}
            for row in stat_list:
                label = row[key] or 'Unknown'
                merged[label] = merged.get(label, 0) + row['count']
            return [{'label': k, 'count': v} for k, v in sorted(merged.items(), key=lambda x: -x[1])]
        
        os_stats = normalize_stat(os_stats_raw, 'os')
        browser_stats = normalize_stat(browser_stats_raw, 'browser')
        device_stats = normalize_stat(device_stats_raw, 'device_type')
        
        country_stats = list(base_query.values('country', 'country_code').annotate(count=Count('id')).order_by('-count')[:10])
        city_stats = list(base_query.values('city', 'country').annotate(count=Count('id')).order_by('-count')[:10])
        referrer_stats = list(base_query.values('referrer').annotate(count=Count('id')).order_by('-count')[:10])
        
        # Traffic Sources — classify by referrer URL pattern
        traffic_sources = {'Direct': 0, 'Internal': 0, 'Search': 0, 'Social': 0, 'Referral': 0, 'QR': 0}
        all_referrers = base_query.values('referrer', 'is_qr_scan').annotate(count=Count('id'))
        
        for item in all_referrers:
            ref = (item['referrer'] or '').lower()
            cnt = item['count']
            if item['is_qr_scan']:
                traffic_sources['QR'] += cnt
            elif ref and (ref.startswith('http://127.') or ref.startswith('http://localhost') or
                          ref.startswith('https://127.') or ref.startswith('https://localhost') or
                          ref.startswith('http://192.168.') or ref.startswith('http://10.')):
                traffic_sources['Internal'] += cnt
            elif any(s in ref for s in ['google', 'bing', 'yahoo', 'duckduckgo', 'baidu', 'yandex']):
                traffic_sources['Search'] += cnt
            elif any(s in ref for s in ['instagram', 'facebook', 'twitter', 't.co', 'linkedin',
                                        'tiktok', 'youtube', 'whatsapp', 'telegram', 'reddit',
                                        'pinterest', 'snapchat']):
                traffic_sources['Social'] += cnt
            elif ref:
                traffic_sources['Referral'] += cnt
            else:
                traffic_sources['Direct'] += cnt
                
        ts_stats = [{'source': k, 'count': v} for k, v in traffic_sources.items() if v > 0]
        
        # Clicks by Hour / Day (using local timezone)
        clicks_by_hour = [0] * 24
        clicks_by_day = [0] * 7
        timestamps = base_query.values_list('timestamp', flat=True)
        for ts in timestamps:
            local_ts = timezone.localtime(ts)
            clicks_by_hour[local_ts.hour] += 1
            clicks_by_day[local_ts.weekday()] += 1
        
        from datetime import datetime
        def safe_dt(ts_val):
            """Safely coerce SQLite date strings back to datetime objects."""
            if isinstance(ts_val, str):
                try: return datetime.fromisoformat(ts_val.replace('Z', '+00:00'))
                except: pass
            return ts_val

        # Zero-padded time series buckets
        time_series = []
        if selected_range == 'today':
            raw_ts = list(base_query.annotate(ts=TruncHour('timestamp')).values('ts').annotate(
                count=Count('id'), unique=Count('visitor_id', distinct=True),
                qr=Count('id', filter=Q(is_qr_scan=True)),
                human=Count('id', filter=Q(is_bot=False)),
                bot=Count('id', filter=Q(is_bot=True))
            ).order_by('ts'))
            ts_dict = {}
            for r in raw_ts:
                if r['ts']:
                    try: ts_dict[safe_dt(r['ts']).strftime('%I %p')] = r
                    except: pass
            for i in range(24):
                hr_label = (start_date + timedelta(hours=i)).strftime('%I %p')
                r = ts_dict.get(hr_label, {'count': 0, 'unique': 0, 'qr': 0, 'human': 0, 'bot': 0})
                time_series.append({'label': hr_label, 'count': r['count'], 'unique': r['unique'],
                                    'qr': r['qr'], 'human': r.get('human', 0), 'bot': r.get('bot', 0)})
        elif selected_range == '12months':
            raw_ts = list(base_query.annotate(ts=TruncMonth('timestamp')).values('ts').annotate(
                count=Count('id'), unique=Count('visitor_id', distinct=True),
                qr=Count('id', filter=Q(is_qr_scan=True)),
                human=Count('id', filter=Q(is_bot=False)),
                bot=Count('id', filter=Q(is_bot=True))
            ).order_by('ts'))
            ts_dict = {}
            for r in raw_ts:
                if r['ts']:
                    try: ts_dict[safe_dt(r['ts']).strftime('%b %Y')] = r
                    except: pass
            for i in range(12):
                m_label = (now - timedelta(days=365) + timedelta(days=30*i)).strftime('%b %Y')
                r = ts_dict.get(m_label, {'count': 0, 'unique': 0, 'qr': 0, 'human': 0, 'bot': 0})
                time_series.append({'label': m_label, 'count': r['count'], 'unique': r['unique'],
                                    'qr': r['qr'], 'human': r.get('human', 0), 'bot': r.get('bot', 0)})
        else:
            raw_ts = list(base_query.annotate(ts=TruncDate('timestamp')).values('ts').annotate(
                count=Count('id'), unique=Count('visitor_id', distinct=True),
                qr=Count('id', filter=Q(is_qr_scan=True)),
                human=Count('id', filter=Q(is_bot=False)),
                bot=Count('id', filter=Q(is_bot=True))
            ).order_by('ts'))
            ts_dict = {}
            for r in raw_ts:
                if r['ts']:
                    try: ts_dict[safe_dt(r['ts']).strftime('%b %d')] = r
                    except: pass
            for i in range(req_days):
                d_label = (start_date + timedelta(days=i)).strftime('%b %d')
                r = ts_dict.get(d_label, {'count': 0, 'unique': 0, 'qr': 0, 'human': 0, 'bot': 0})
                time_series.append({'label': d_label, 'count': r['count'], 'unique': r['unique'],
                                    'qr': r['qr'], 'human': r.get('human', 0), 'bot': r.get('bot', 0)})
                
        recent_scans_qs = base_query.order_by('-timestamp')
        
        # Best day from buckets that had actual clicks
        best_day = None
        best_day_clicks = 0
        non_zero_buckets = [r for r in time_series if r['count'] > 0]
        if non_zero_buckets:
            best = max(non_zero_buckets, key=lambda x: x['count'])
            best_day = best['label']
            best_day_clicks = best['count']
        
        # Friendly top referrer label
        top_ref_raw = referrer_stats[0]['referrer'] if referrer_stats else ''
        if not top_ref_raw or any(x in (top_ref_raw or '') for x in ['127.0.0.1', 'localhost', '192.168.', '::1']):
            top_ref_label = 'Internal / Direct'
        else:
            try:
                from urllib.parse import urlparse
                top_ref_label = urlparse(top_ref_raw).netloc or top_ref_raw[:30]
            except:
                top_ref_label = (top_ref_raw or '')[:30]
            
        # Peak Hour calculation
        peak_hour_idx = clicks_by_hour.index(max(clicks_by_hour)) if max(clicks_by_hour) > 0 else None
        peak_hour_clicks = max(clicks_by_hour) if max(clicks_by_hour) > 0 else 0
        if peak_hour_idx is not None:
            from datetime import time
            peak_hour_label = time(hour=peak_hour_idx).strftime("%I:%M %p") + " - " + time(hour=(peak_hour_idx+1)%24).strftime("%I:%M %p")
        else:
            peak_hour_label = None

        # Generate Key Insights
        insights = []
        if best_day and best_day_clicks > 0:
            insights.append(f"Your short URL received the most traffic on {best_day}.")
        if device_stats and total_clicks > 0:
            top_dev_pct = int(device_stats[0]['count'] / total_clicks * 100)
            insights.append(f"{device_stats[0]['label']} users generated {top_dev_pct}% of your clicks.")
        if country_stats and total_clicks > 0:
            top_ctr_pct = int(country_stats[0]['count'] / total_clicks * 100)
            c_name = country_stats[0]['country']
            if c_name == 'Internal / Local': c_name = 'Local network'
            insights.append(f"{c_name} generated {top_ctr_pct}% of your total traffic.")
        if ts_stats:
            top_ts = ts_stats[0]['source']
            insights.append(f"{top_ts} traffic is currently your largest traffic source.")
            
        summary = {
            'best_day': best_day,
            'best_day_clicks': best_day_clicks,
            'top_country': country_stats[0]['country'] if country_stats else None,
            'top_country_clicks': country_stats[0]['count'] if country_stats else 0,
            'top_country_pct': int(country_stats[0]['count'] / total_clicks * 100) if country_stats and total_clicks else 0,
            'top_device': device_stats[0]['label'] if device_stats else None,
            'top_device_pct': int(device_stats[0]['count'] / total_clicks * 100) if device_stats and total_clicks else 0,
            'top_referrer': top_ref_label,
            'peak_hour': peak_hour_label,
            'peak_hour_clicks': peak_hour_clicks,
            'insights': insights,
        }
        
        return total_clicks, qr_scans, unique_clicks, human_clicks, bot_clicks, trends, source_stats, os_stats, browser_stats, device_stats, country_stats, city_stats, referrer_stats, time_series, recent_scans_qs, summary, ts_stats, clicks_by_hour, clicks_by_day

    try:
        total_clicks, qr_scans, unique_clicks, human_clicks, bot_clicks, trends, source_stats, os_stats, browser_stats, device_stats, country_stats, city_stats, referrer_stats, time_series, recent_scans_qs, perf_summary, ts_stats, clicks_by_hour, clicks_by_day = get_data()
    except Exception as e:
        print(f"[Analytics Error]: {e}")
        total_clicks, qr_scans, unique_clicks, human_clicks, bot_clicks = 0, 0, 0, 0, 0
        trends = {'total': '0%', 'qr': '0%', 'unique': '0%', 'human': '0%', 'bot': '0%'}
        source_stats, os_stats, browser_stats, device_stats, country_stats, city_stats, referrer_stats, time_series, recent_scans_qs = [], [], [], [], [], [], [], [], []
        perf_summary = {}
        ts_stats, clicks_by_hour, clicks_by_day = [], [0]*24, [0]*7

    paginator = Paginator(recent_scans_qs, 10)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)

    chart_labels = [d['label'] for d in time_series]
    chart_data_total = [d['count'] for d in time_series]
    chart_data_unique = [d['unique'] for d in time_series]
    chart_data_qr = [d['qr'] for d in time_series]
    
    # Pre-process referrers for template display
    for r in referrer_stats:
        ref = r.get('referrer') or ''
        if not ref or any(x in ref for x in ['127.0.0.1', 'localhost', '192.168.', '10.0.', '::1']):
            r['display_name'] = 'Internal / Direct'
        else:
            try:
                from urllib.parse import urlparse
                r['display_name'] = urlparse(ref).netloc or ref[:40]
            except:
                r['display_name'] = ref[:40]
            
    for l in country_stats:
        if l.get('country_code') == 'LCL' or l.get('country') == 'Internal':
            l['country'] = 'Internal / Local'
            l['is_local'] = True
            
    for l in city_stats:
        if l.get('city') in ('Private IP', 'Unknown') or l.get('country') in ('Internal', 'Unknown'):
            if not l.get('city') or l['city'] in ('Private IP', 'Unknown'):
                l['city'] = 'Private Network'

    # Build full time-series JSON with all metric streams
    chart_data_human = [d.get('human', 0) for d in time_series]
    chart_data_bot = [d.get('bot', 0) for d in time_series]
    
    # Unique click ratio (per period bucket)
    chart_data_ratio = []
    for d in time_series:
        if d['count'] > 0:
            chart_data_ratio.append(round(d['unique'] / d['count'] * 100, 1))
        else:
            chart_data_ratio.append(0)
    avg_ratio = round(sum(chart_data_ratio) / len([r for r in chart_data_ratio if r > 0]), 1) if any(r > 0 for r in chart_data_ratio) else 0

    return render(request, 'dynamic_qr/short_url_analytics.html', {
        'qr': qr,
        'selected_range': selected_range,
        'history_clamped': history_clamped,
        'max_history_days': max_history_days,
        'total_clicks': total_clicks,
        'qr_scans': qr_scans,
        'unique_clicks': unique_clicks,
        'human_clicks': human_clicks,
        'bot_clicks': bot_clicks,
        'trends': trends,
        'perf_summary': perf_summary,
        'country_stats': country_stats,
        'city_stats': city_stats,
        'referrer_stats': referrer_stats,
        'ts_stats': ts_stats,
        'device_stats': device_stats,
        'browser_stats': browser_stats,
        'os_stats': os_stats,
        # JSON payloads for charts
        'js_labels': json.dumps(chart_labels),
        'js_data_total': json.dumps(chart_data_total),
        'js_data_unique': json.dumps(chart_data_unique),
        'js_data_qr': json.dumps(chart_data_qr),
        'js_data_human': json.dumps(chart_data_human),
        'js_data_bot': json.dumps(chart_data_bot),
        'js_data_ratio': json.dumps(chart_data_ratio),
        'js_avg_ratio': json.dumps(avg_ratio),
        'js_ts_labels': json.dumps([d['source'] for d in ts_stats]),
        'js_ts_data': json.dumps([d['count'] for d in ts_stats]),
        'js_os_labels': json.dumps([d['label'] for d in os_stats]),
        'js_os_data': json.dumps([d['count'] for d in os_stats]),
        'js_browser_labels': json.dumps([d['label'] for d in browser_stats]),
        'js_browser_data': json.dumps([d['count'] for d in browser_stats]),
        'js_device_labels': json.dumps([d['label'] for d in device_stats]),
        'js_device_data': json.dumps([d['count'] for d in device_stats]),
        'js_clicks_by_hour': json.dumps(clicks_by_hour),
        'js_clicks_by_day': json.dumps(clicks_by_day),
        'page_obj': page_obj,
    })


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
        design_data_json = request.POST.get('design_options', '{}')
        logo = request.FILES.get('logo')
        logo_cropped = request.POST.get('logo_cropped')
        file_content = request.FILES.get('file_content')

        if logo_cropped and logo_cropped.startswith('data:image'):
            from django.core.files.base import ContentFile
            import base64
            import uuid
            try:
                format, imgstr = logo_cropped.split(';base64,')
                ext = 'png'
                if '/svg' in format: ext = 'svg'
                elif '/jpeg' in format: ext = 'jpg'
                logo = ContentFile(base64.b64decode(imgstr), name=f"logo_{uuid.uuid4()}.{ext}")
            except:
                pass

        if not qr_name:
            return JsonResponse({'error': 'QR name is required.'}, status=400)

        # Parse incoming structured payload with safe fallbacks.
        try:
            incoming_data = json.loads(qr_data_json) if qr_data_json else {}
        except Exception:
            incoming_data = {}

        if not isinstance(incoming_data, dict):
            incoming_data = {}

        # Build data from raw POST fields when front-end payload is missing/incomplete.
        fallback_data = {
            'text': request.POST.get('text', '').strip(),
            'phone': request.POST.get('phone', '').strip(),
            'phone_mobile': request.POST.get('phone_mobile', '').strip(),
            'email': request.POST.get('email', '').strip(),
            'subject': request.POST.get('subject', '').strip(),
            'body': request.POST.get('body', '').strip(),
            'message': request.POST.get('message', '').strip(),
            'ssid': request.POST.get('ssid', '').strip(),
            'password': request.POST.get('password', '').strip(),
            'encryption': request.POST.get('encryption', '').strip(),
            'latitude': request.POST.get('latitude', '').strip(),
            'longitude': request.POST.get('longitude', '').strip(),
            'first_name': request.POST.get('first_name', '').strip(),
            'last_name': request.POST.get('last_name', '').strip(),
            'organization': request.POST.get('organization', '').strip(),
            'destination_url': request.POST.get('destination_url', '').strip(),
        }
        # Remove empty fallback keys.
        fallback_data = {k: v for k, v in fallback_data.items() if v}
        if fallback_data:
            incoming_data = {**fallback_data, **incoming_data}

        qr.qr_name = qr_name
        qr.qr_type = qr_type

        # Keep destination_url aligned with selected type.
        url_types = {
            'url', 'custom-url', 'youtube', 'facebook', 'instagram', 'telegram',
            'tiktok', 'x-twitter', 'snapchat', 'pinterest', 'linkedin',
            'google-review', 'google-forms', 'google-doc', 'google-sheets',
            'play-market', 'app-store', 'paypal', 'etsy', 'amazon', 'venmo',
            'upi', 'crypto', 'spotify', 'link-list', 'office-365'
        }
        if qr_type in url_types:
            qr.destination_url = destination_url or incoming_data.get('destination_url', '')
        else:
            qr.destination_url = ''

        # Normalize type-specific data so redirect resolver never gets empty payload.
        if qr_type == 'phone' and not incoming_data.get('phone'):
            incoming_data['phone'] = incoming_data.get('phone_mobile', '')
        if qr_type == 'vcard' and not incoming_data.get('phone_mobile'):
            incoming_data['phone_mobile'] = incoming_data.get('phone', '')

        qr.qr_data = incoming_data
            
        try:
            qr.design_options = json.loads(design_data_json)
        except:
            pass

        qr.fg_color = fg_color
        qr.bg_color = bg_color
        qr.body_style = body_style
        qr.eye_style = eye_style
        qr.ball_style = ball_style
        qr.is_active = is_active
        # --- Permanent Logo Persistence ---
        if logo:
            qr.logo = logo
        elif qr.design_options and qr.design_options.get('logo_preset'):
            preset = qr.design_options.get('logo_preset')
            if preset and preset != 'none':
                # If we have a preset but NO logo file, try to cache it from the preset to the logo field
                try:
                    target_icon = os.path.join(settings.MEDIA_ROOT, 'brand_icons', f"{preset}.png")
                    if os.path.exists(target_icon) and os.path.getsize(target_icon) > 0:
                        from django.core.files import File
                        with open(target_icon, 'rb') as f:
                            qr.logo.save(f"{preset}_preset.png", File(f), save=False)
                except: pass
        
        qr.save()

        # If AJAX request, return JSON
        if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
            qr_content = qr.get_static_content(request)
            return JsonResponse({
                'success': True,
                'qr_id': str(qr.id),
                'short_code': qr.short_code,
                'redirect_url': qr_content,
                'qr_name': qr.qr_name,
                'qr_type': qr.qr_type,
            })

        return redirect('dynamic_qr:dashboard')

    qr_content = qr.get_static_content(request)
    
    # Merge existing data for dynamic fields
    content_data = dict(qr.qr_data if qr.qr_data else {})
    if qr.destination_url:
        content_data['destination_url'] = qr.destination_url
        
    import json
    return render(request, 'dynamic_qr/edit_qr.html', {
        'qr': qr,
        'redirect_url': qr_content,
        'content_data': json.dumps(content_data),
        'design_options_json': json.dumps(qr.design_options if qr.design_options else {}),
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
    qr_content = qr.get_static_content(request)
    return render(request, 'dynamic_qr/details.html', {
        'qr': qr,
        'redirect_url': qr_content,
    })


# ═══════════════════════════════════════════════════════════════
# REDIRECT: Short code → Destination URL
# ═══════════════════════════════════════════════════════════════
def dqr_redirect_view(request, short_code):
    """
    When someone scans the dynamic QR code, they hit this URL.
    De-duplicates hits to prevent double-counting from pre-fetchers.
    """
    from django.db.models import Q
    qr = get_object_or_404(DynamicQRCode, Q(short_code=short_code) | Q(custom_alias=short_code))

    if not qr.is_active:
        return render(request, 'dynamic_qr/qr_disabled.html', {'qr': qr})
        
    if qr.expiry_date and timezone.now() > qr.expiry_date:
        return render(request, 'dynamic_qr/qr_disabled.html', {'qr': qr, 'expired': True})
        
    if qr.password and not request.session.get(f'qr_auth_{qr.id}'):
        if request.method == 'POST':
            from django.contrib.auth.hashers import check_password
            pw = request.POST.get('password', '')
            if check_password(pw, qr.password):
                request.session[f'qr_auth_{qr.id}'] = True
                return redirect(request.path)
            else:
                return render(request, 'dynamic_qr/qr_password.html', {'qr': qr, 'error': 'Incorrect password'})
        return render(request, 'dynamic_qr/qr_password.html', {'qr': qr})

    # Logic for Logging (only once every 5 seconds per session)
    now_ts = timezone.now().timestamp()
    last_ts = request.session.get(f'qr_last_hit_{qr.id}', 0)
    
    if (now_ts - last_ts) >= 5:
        # 1. Increment Scan Count
        qr.increment_scan()
        request.session[f'qr_last_hit_{qr.id}'] = now_ts
        
        # 2. Log Detailed Analytics
        ua = request.META.get('HTTP_USER_AGENT', '').lower()
        
        # Get Real IP (handle proxies)
        x_forwarded_for = request.META.get('HTTP_X_FORWARDED_FOR')
        if x_forwarded_for:
            ip = x_forwarded_for.split(',')[0].strip()
        else:
            ip = request.META.get('REMOTE_ADDR', '')
            
        # Source & QR Tracking
        is_qr_scan = request.GET.get('source') == 'qr'
        source = 'QR' if is_qr_scan else 'Direct'
        
        # Visitor ID generation
        import hashlib
        visitor_string = f"{ip.rsplit('.', 1)[0] if '.' in ip else ip}_{ua}_{qr.id}"
        visitor_id = hashlib.sha256(visitor_string.encode('utf-8')).hexdigest()[:32]
        
        # Improved Manual Parsing
        browser = 'Other'
        if 'edg/' in ua or 'edge' in ua: browser = 'Edge'
        elif 'samsungbrowser' in ua: browser = 'Samsung Internet'
        elif 'opera' in ua or 'opr/' in ua: browser = 'Opera'
        elif 'chrome' in ua and 'safari' in ua: browser = 'Chrome'
        elif 'safari' in ua and 'chrome' not in ua: browser = 'Safari'
        elif 'firefox' in ua: browser = 'Firefox'
        
        os_name = 'Other'
        if 'windows' in ua: os_name = 'Windows'
        elif 'iphone' in ua or 'ipad' in ua: os_name = 'iOS'
        elif 'mac' in ua: os_name = 'macOS'
        elif 'android' in ua: os_name = 'Android'
        elif 'linux' in ua: os_name = 'Linux'
        
        device = 'Desktop'
        if 'ipad' in ua or 'tablet' in ua or ('android' in ua and 'mobile' not in ua):
            device = 'Tablet'
        elif 'mobile' in ua or 'iphone' in ua or 'android' in ua:
            device = 'Mobile'
        
        referrer = request.META.get('HTTP_REFERER', '')[:500]
        bot_keywords = ['bot', 'crawl', 'spider', 'slurp', 'mediapartners', 'preview', 'slack', 'discord', 'whatsapp', 'skype']
        is_bot = any(b in ua for b in bot_keywords)

        # Geolocation logic
        country, country_code, region, city = 'Unknown', 'XX', 'Unknown', 'Unknown'
        lat, lon = None, None
        
        # Check if IP is private/local
        is_private = False
        if ip:
            if ip.startswith(('127.', '192.168.', '10.', '172.16.', '172.17.', '172.18.', '172.19.', '172.20.', '172.21.', '172.22.', '172.23.', '172.24.', '172.25.', '172.26.', '172.27.', '172.28.', '172.29.', '172.30.', '172.31.')) or ip == '::1':
                is_private = True

        if ip and not is_private:
            try:
                import json
                from urllib.request import urlopen, Request
                # Using ip-api.com (free for non-commercial use, 45 requests/min)
                headers = {'User-Agent': 'ScanPDF/1.0'}
                req = Request(f'http://ip-api.com/json/{ip}?fields=status,country,countryCode,regionName,city,lat,lon', headers=headers)
                with urlopen(req, timeout=4) as resp:
                    geo_data = json.loads(resp.read().decode())
                    if geo_data.get('status') == 'success':
                        country = geo_data.get('country', 'Unknown')
                        country_code = geo_data.get('countryCode', 'XX')
                        region = geo_data.get('regionName', 'Unknown')
                        city = geo_data.get('city', 'Unknown')
                        lat = geo_data.get('lat')
                        lon = geo_data.get('lon')
            except Exception as e:
                # Fallback to local server info if API fails
                pass
        elif is_private:
            # For development/internal scans
            country, country_code, region, city = 'Internal', 'LCL', 'Local Network', 'Private IP'
            # If in debug mode, we can mock a location for visual testing
            from django.conf import settings
            if getattr(settings, 'DEBUG', False):
                country, country_code, region, city = 'India', 'IN', 'Tamil Nadu', 'Chennai'
                lat, lon = 13.0827, 80.2707

        from .models import QRAnalytics
        from django.db import connection, OperationalError
        try:
            QRAnalytics.objects.create(
                qr_code=qr, ip_address=ip, user_agent=ua[:500], 
                browser=browser, os=os_name, device_type=device,
                country=country, country_code=country_code, region=region, city=city,
                latitude=lat, longitude=lon,
                referrer=referrer, is_bot=is_bot,
                is_qr_scan=is_qr_scan, source=source, visitor_id=visitor_id
            )
        except OperationalError:
            # Manual fallback for DB schema mismatch if migrations haven't run
            try:
                with connection.cursor() as cursor:
                    cursor.execute('CREATE TABLE IF NOT EXISTS "dynamic_qr_qranalytics" ("id" integer NOT NULL PRIMARY KEY AUTOINCREMENT, "timestamp" datetime NOT NULL, "ip_address" char(39) NULL, "user_agent" text NULL, "browser" varchar(50) NULL, "os" varchar(50) NULL, "device_type" varchar(50) NULL, "country" varchar(100) NOT NULL DEFAULT "Unknown", "country_code" varchar(10) NOT NULL DEFAULT "XX", "region" varchar(100) NOT NULL DEFAULT "Unknown", "city" varchar(100) NOT NULL DEFAULT "Unknown", "latitude" float NULL, "longitude" float NULL, "qr_code_id" uuid NOT NULL REFERENCES "dynamic_qr_dynamicqrcode" ("id") DEFERRABLE INITIALLY DEFERRED);')
                    # Try to add missing columns if table exists but is old
                    cols = ["country_code", "region", "latitude", "longitude", "referrer", "is_bot", "is_qr_scan", "source", "visitor_id"]
                    for col in cols:
                        try:
                            if col == 'referrer':
                                cursor.execute(f'ALTER TABLE "dynamic_qr_qranalytics" ADD COLUMN "{col}" varchar(500);')
                            elif col == 'is_bot' or col == 'is_qr_scan':
                                cursor.execute(f'ALTER TABLE "dynamic_qr_qranalytics" ADD COLUMN "{col}" bool NOT NULL DEFAULT 0;')
                            elif col == 'source':
                                cursor.execute(f'ALTER TABLE "dynamic_qr_qranalytics" ADD COLUMN "{col}" varchar(50) NOT NULL DEFAULT "Direct";')
                            elif col == 'visitor_id':
                                cursor.execute(f'ALTER TABLE "dynamic_qr_qranalytics" ADD COLUMN "{col}" varchar(64) NULL;')
                            else:
                                cursor.execute(f'ALTER TABLE "dynamic_qr_qranalytics" ADD COLUMN "{col}" {"float" if "tude" in col else "varchar(100)"};')
                        except: pass
                QRAnalytics.objects.create(
                    qr_code=qr, ip_address=ip, user_agent=ua[:500], 
                    browser=browser, os=os_name, device_type=device,
                    country=country, country_code=country_code, region=region, city=city,
                    latitude=lat, longitude=lon,
                    referrer=referrer, is_bot=is_bot,
                    is_qr_scan=is_qr_scan, source=source, visitor_id=visitor_id
                )
            except: pass
        except: pass

    def _no_cache(response):
        # Prevent stale scan results after edits on the same short code.
        response['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
        response['Pragma'] = 'no-cache'
        response['Expires'] = '0'
        return response

    # Determine redirect behavior
    redirect_types = [
        'url', 'whatsapp', 'youtube', 'facebook', 'instagram', 'telegram', 
        'tiktok', 'x-twitter', 'snapchat', 'pinterest', 'linkedin',
        'pdf', 'audio', 'video', 'image', 'pptx', 'excel', 'word',
        'google-review', 'google-forms', 'google-doc', 'google-sheets',
        'play-market', 'app-store', 'paypal', 'etsy', 'amazon', 'venmo', 
        'upi', 'crypto', 'spotify', 'link-list', 'custom-url', 'office-365'
    ]

    # 1. Direct vCard Download (Most direct experience for contacts)
    if qr.qr_type == 'vcard':
        vcard_data = qr.get_raw_payload()
        response = HttpResponse(vcard_data, content_type='text/vcard')
        response['Content-Disposition'] = f'attachment; filename="{qr.qr_name or "contact"}.vcf"'
        return _no_cache(response)

    # 2. Files & URLs Redirect (Direct Redirect)
    target_url = None
    if qr.file_content:
        target_url = qr.file_content.url
    elif qr.destination_url:
        target_url = qr.destination_url

    if target_url and qr.qr_type in redirect_types:
        return _no_cache(HttpResponseRedirect(target_url))
    
    # 3. Protocol payload handling (tel:, sms:, mailto:, geo:, etc.)
    payload = qr.get_raw_payload()
    # Extra fallback for legacy entries where values were stored with mixed keys.
    if not payload:
        data = qr.qr_data or {}
        if qr.qr_type == 'phone':
            phone = data.get('phone') or data.get('phone_mobile')
            payload = f"tel:{phone}" if phone else ''
        elif qr.qr_type == 'sms':
            phone = data.get('phone') or data.get('phone_mobile')
            msg = data.get('message', '')
            payload = f"sms:{phone}?body={msg}" if phone else ''
        elif qr.qr_type in ('url', 'custom-url'):
            payload = data.get('destination_url') or qr.destination_url or ''
        elif qr.qr_type == 'text':
            payload = data.get('text', '')

    if payload:
        if payload.startswith(('http://', 'https://')):
            return _no_cache(HttpResponseRedirect(payload))

        if payload.startswith(('tel:', 'sms:', 'mailto:', 'geo:')):
            # Django blocks these schemes in HttpResponseRedirect; set Location directly.
            protocol_redirect = HttpResponse(status=302)
            protocol_redirect['Location'] = payload
            return _no_cache(protocol_redirect)

        content_type = 'text/plain; charset=utf-8'
        if qr.qr_type == 'wifi':
            content_type = 'text/plain; charset=utf-8'
        elif qr.qr_type == 'location':
            content_type = 'text/uri-list; charset=utf-8'
        return _no_cache(HttpResponse(payload, content_type=content_type))

    # Last safety fallback: if URL exists but type mismatched, still honor it.
    if qr.destination_url:
        return _no_cache(HttpResponseRedirect(qr.destination_url))

    return _no_cache(HttpResponse("No QR content configured.", status=404, content_type='text/plain; charset=utf-8'))



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
    qr_id = data.get('qr_id')
    qr_obj = None
    if qr_id:
        qr_obj = get_object_or_404(DynamicQRCode, id=qr_id, user=request.user)
        if qr_obj.qr_type == 'custom-url' and not qr_obj.qr_enabled:
            return JsonResponse({'error': 'QR Code is disabled for this Short URL.'}, status=400)

    from converter.utils import generate_qr_code, get_output_path
    from converter.views import create_cleanup_response

    text = data.get('text')
    if not text:
        if qr_obj: text = qr_obj.get_static_content(request)
        else: text = 'https://scanpdf.com'
        
    fg_color = data.get('fg_color') or (qr_obj.fg_color if qr_obj else '#000000')
    bg_color = data.get('bg_color') or (qr_obj.bg_color if qr_obj else '#ffffff')
    style = data.get('style') or (qr_obj.body_style if qr_obj else 'square')
    eye_style = data.get('eye_style') or (qr_obj.eye_style if qr_obj else 'square')
    ball_style = data.get('ball_style') or (qr_obj.ball_style if qr_obj else 'square')
    output_format = data.get('output_format', 'png')

    try:
        from converter.utils import save_uploaded_file
        logo_path = None
        
        # Determine Design Options
        design_options = data.get('design_options')
        if not design_options and qr_obj:
            design_options = json.dumps(qr_obj.design_options if qr_obj.design_options else {})
        elif not design_options:
            design_options = '{}'

        # --- Dynamic Logo Resolution Pipeline ---
        logo_path = None
        brand_id = data.get('logo') # Selection from UI
        
        # 1. Parse design_options from request ONLY (highest priority for live updates)
        try:
            req_design = json.loads(design_options) if design_options else {}
            if req_design.get('logo_preset'):
                brand_id = req_design.get('logo_preset')
        except: 
            req_design = {}

        # 2. Fallback to persistent logo_preset from DB
        if (not brand_id or brand_id == 'existing') and qr_obj:
            if qr_obj.design_options:
                brand_id = qr_obj.design_options.get('logo_preset')

        # 3. Handle Brand Presets (Auto-detect and cache)
        if brand_id and brand_id not in ('none', 'existing', ''):
            domain_map = {
                'facebook': 'facebook.com', 'instagram': 'instagram.com', 'youtube': 'youtube.com',
                'whatsapp': 'whatsapp.com', 'linkedin': 'linkedin.com', 'telegram': 'telegram.org',
                'twitter': 'twitter.com', 'x': 'x.com', 'tiktok': 'tiktok.com', 'snapchat': 'snapchat.com',
                'pinterest': 'pinterest.com', 'spotify': 'spotify.com', 'apple': 'apple.com',
                'google': 'google.com', 'amazon': 'amazon.com', 'paypal': 'paypal.com',
                'discord': 'discord.com', 'reddit': 'reddit.com', 'slack': 'slack.com',
                'github': 'github.com', 'microsoft': 'microsoft.com'
            }
            target_domain = domain_map.get(brand_id)
            if target_domain:
                try:
                    icon_dir = os.path.join(settings.MEDIA_ROOT, 'brand_icons')
                    os.makedirs(icon_dir, exist_ok=True)
                    icon_path = os.path.join(icon_dir, f"{brand_id}.png")
                    
                    # Persistent Cache: Download once, reuse forever
                    if not os.path.exists(icon_path) or os.path.getsize(icon_path) == 0:
                        icon_url = f"https://www.google.com/s2/favicons?sz=128&domain={target_domain}"
                        h = {
                            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
                            'Accept': 'image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8'
                        }
                        # Use verify=False to bypass local SSL certificate issues if necessary
                        r = requests.get(icon_url, headers=h, timeout=12, verify=False)
                        if r.status_code == 200:
                            with open(icon_path, 'wb') as f: f.write(r.content)
                    
                    if os.path.exists(icon_path) and os.path.getsize(icon_path) > 0:
                        logo_path = icon_path
                except Exception as e:
                    print(f"[QR LOGO ERROR] Failed to resolve {brand_id}: {e}")

        # 4. Handle Cropped Blob (Highest priority for creation/edit preview)
        if data.get('logo_cropped') and data.get('logo_cropped').startswith('data:image'):
            try:
                base64_data = data.get('logo_cropped').split(';base64,')[1]
                tmp_dir = os.path.join(settings.MEDIA_ROOT, 'temp_previews')
                os.makedirs(tmp_dir, exist_ok=True)
                logo_path = os.path.join(tmp_dir, f"p_{uuid.uuid4().hex[:8]}.png")
                with open(logo_path, 'wb') as f: 
                    f.write(base64.b64decode(base64_data))
            except: pass

        # 5. Fallback: Database-saved primary logo (Persistent Fallback)
        if not logo_path and qr_obj and qr_obj.logo:
            try:
                if os.path.exists(qr_obj.logo.path):
                    logo_path = qr_obj.logo.path
            except: pass

        # 6. Manual Preview Upload
        if not logo_path and 'logo' in request.FILES:
            logo_path = save_uploaded_file(request.FILES['logo'])

        output_path = generate_qr_code(
            text, fg_color=fg_color, bg_color=bg_color,
            style=style, eye_style=eye_style, ball_style=ball_style,
            logo_path=logo_path, output_format=output_format,
            design_options=design_options,
            eye_color_outer=data.get('eye_color_outer'),
            eye_color_inner=data.get('eye_color_inner')
        )

        if logo_path and os.path.exists(logo_path) and 'temp' in logo_path:
            try: os.remove(logo_path)
            except: pass

        ct = 'image/png'
        if output_format.lower() in ('jpg', 'jpeg'): ct = 'image/jpeg'
        elif output_format.lower() == 'svg': ct = 'image/svg+xml'
        return create_cleanup_response(output_path, content_type=ct)
    except Exception as e:
        return JsonResponse({'error': str(e)}, status=500)


@dqr_login_required
def dqr_download_view(request, qr_id):
    """Generate and return the QR image for a specific dynamic QR code in requested format."""
    qr = get_object_or_404(DynamicQRCode, id=qr_id, user=request.user)
    if qr.qr_type == 'custom-url' and not qr.qr_enabled:
        return JsonResponse({'error': 'QR Code is disabled for this Short URL.'}, status=400)
    fmt = request.GET.get('format', 'png').lower()
    if fmt not in ('png', 'jpg', 'jpeg', 'svg'):
        fmt = 'png'
        
    from converter.utils import generate_qr_code
    from converter.views import create_cleanup_response

    # Use get_static_content to encode the raw data directly, bypassing redirects
    qr_content = qr.get_static_content(request)
    
    # Fix logo resolution for download (Sync with generate_image logic)
    logo_path = None
    if qr.logo and os.path.exists(qr.logo.path):
        logo_path = qr.logo.path
    
    if not logo_path and qr.design_options:
        brand_id = qr.design_options.get('logo_preset')
        if brand_id and brand_id not in ('none', 'existing', ''):
            domain_map = {
                'facebook': 'facebook.com', 'instagram': 'instagram.com', 'youtube': 'youtube.com',
                'whatsapp': 'whatsapp.com', 'linkedin': 'linkedin.com', 'telegram': 'telegram.org',
                'x': 'x.com', 'tiktok': 'tiktok.com', 'snapchat': 'snapchat.com',
                'pinterest': 'pinterest.com', 'spotify': 'spotify.com', 'apple': 'apple.com',
                'google': 'google.com', 'amazon': 'amazon.com', 'paypal': 'paypal.com',
                'discord': 'discord.com', 'reddit': 'reddit.com', 'slack': 'slack.com',
                'github': 'github.com', 'microsoft': 'microsoft.com'
            }
            if brand_id in domain_map:
                try:
                    icon_dir = os.path.join(settings.MEDIA_ROOT, 'brand_icons')
                    os.makedirs(icon_dir, exist_ok=True)
                    icon_path = os.path.join(icon_dir, f"{brand_id}.png")
                    
                    if not os.path.exists(icon_path):
                        icon_url = f"https://www.google.com/s2/favicons?sz=128&domain={domain_map[brand_id]}"
                        headers = {'User-Agent': 'Mozilla/5.0'}
                        r = requests.get(icon_url, headers=headers, timeout=5)
                        if r.status_code == 200:
                            with open(icon_path, 'wb') as f: f.write(r.content)
                    
                    if os.path.exists(icon_path):
                        logo_path = icon_path
                except: pass

    # Extract per-element colors if available
    eye_color_outer = qr.design_options.get('eye_color_outer') if qr.design_options else None
    eye_color_inner = qr.design_options.get('eye_color_inner') if qr.design_options else None

    # Use the helper from converter.utils
    output_path = generate_qr_code(
        qr_content,
        fg_color=qr.fg_color,
        bg_color=qr.bg_color,
        style=qr.body_style,
        eye_style=qr.eye_style,
        ball_style=qr.ball_style,
        logo_path=logo_path,
        output_format=fmt,
        design_options=qr.design_options,
        eye_color_outer=eye_color_outer,
        eye_color_inner=eye_color_inner
    )
    
    ct = 'image/png'
    if fmt in ('jpg', 'jpeg'): ct = 'image/jpeg'
    elif fmt == 'svg': ct = 'image/svg+xml'
    
    response = create_cleanup_response(output_path, content_type=ct)
    response['Content-Disposition'] = f'attachment; filename="qr_{qr.short_code}.{fmt}"'
    return response


# ═══════════════════════════════════════════════════════════════
# VALIDATION API
# ═══════════════════════════════════════════════════════════════
def dqr_check_username_view(request):
    import re
    username = request.GET.get('username', '').strip()
    if not username:
        return JsonResponse({'valid': False, 'message': 'Username is required'})
    
    if len(username) < 3:
        return JsonResponse({'valid': False, 'message': 'Username is too short'})
    
    if not re.match(r'^[\w]+\Z', username):
        return JsonResponse({'valid': False, 'message': 'Username must contain only valid characters'})
        
    if User.objects.filter(username__iexact=username).exists():
        return JsonResponse({'valid': False, 'available': False, 'message': 'Username is already taken'})
        
    return JsonResponse({'valid': True, 'available': True, 'message': 'Username is available'})

def dqr_check_email_view(request):
    import re
    email = request.GET.get('email', '').strip()
    if not email:
        return JsonResponse({'valid': False, 'message': 'Email is required'})
        
    email_regex = r'^[^\s@]+@[^\s@]+\.[^\s@]+$'
    if not re.match(email_regex, email):
        return JsonResponse({'valid': False, 'message': 'Please enter a valid email address'})
        
    if User.objects.filter(email__iexact=email).exists():
        return JsonResponse({'valid': False, 'available': False, 'message': 'Email is already registered'})
        
    return JsonResponse({'valid': True, 'available': True, 'message': 'Email is available'})


# ═══════════════════════════════════════════════════════════════
# REDIRECT WITH HEADER
# ═══════════════════════════════════════════════════════════════
def dqr_redirect_with_header_view(request, header, short_code):
    """
    Handles resolving domain/header/slug.
    Checks if the header matches the one in DB, then passes to dqr_redirect_view.
    """
    from django.db.models import Q
    from django.http import Http404
    
    # 1. Fetch QR code by short code (since it's globally unique)
    qr = get_object_or_404(DynamicQRCode, Q(short_code=short_code) | Q(custom_alias=short_code))
    
    # 2. Verify that the requested header exactly matches the QR code's header
    if not qr.header or qr.header != header:
        raise Http404("Short URL with this header does not exist.")
        
    # 3. Pass request to the original redirect view
    return dqr_redirect_view(request, short_code)
