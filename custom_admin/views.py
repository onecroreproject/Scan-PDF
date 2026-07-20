from django.shortcuts import render, get_object_or_404, redirect
from django.contrib.auth import logout
from django.contrib.auth.models import User
from django.db.models import Sum, Count
from django.utils import timezone
from datetime import timedelta
from services.models import Subscription, Payment, Plan, ActivityLog
from dynamic_qr.models import DynamicQRCode
from .decorators import superuser_required

@superuser_required
def dashboard_view(request):
    today = timezone.now().date()
    thirty_days_ago = timezone.now() - timedelta(days=30)
    
    context = {
        'total_users': User.objects.count(),
        'today_users': User.objects.filter(date_joined__date=today).count(),
        'active_users': User.objects.filter(is_active=True).count(),
        'inactive_users': User.objects.filter(is_active=False).count(),
        
        'total_payments': Payment.objects.count(),
        'monthly_revenue': Payment.objects.filter(payment_status='Paid', created_at__gte=thirty_days_ago).aggregate(Sum('amount'))['amount__sum'] or 0,
        'yearly_revenue': Payment.objects.filter(payment_status='Paid', created_at__gte=timezone.now() - timedelta(days=365)).aggregate(Sum('amount'))['amount__sum'] or 0,
        
        'total_subscriptions': Subscription.objects.count(),
        'expired_plans': Subscription.objects.filter(status='Expired').count(),
        'pending_renewals': Subscription.objects.filter(status='Active', end_date__lte=timezone.now() + timedelta(days=7)).count(),
        
        'dynamic_qrs_created': DynamicQRCode.objects.exclude(qr_type='custom-url').count(),
        'short_urls_created': DynamicQRCode.objects.filter(qr_type='custom-url').count(),
    }
    return render(request, 'admin_dashboard/dashboard.html', context)

@superuser_required
def users_view(request):
    # Fetch all users and prefetch subscriptions/payments if needed for performance
    users = User.objects.all().order_by('-date_joined')
    user_data = []
    for u in users:
        sub = Subscription.objects.filter(user=u, status='Active').first()
        qrs = DynamicQRCode.objects.filter(user=u).exclude(qr_type='custom-url').count()
        shorts = DynamicQRCode.objects.filter(user=u, qr_type='custom-url').count()
        user_data.append({
            'user': u,
            'sub': sub,
            'qr_count': qrs,
            'short_count': shorts,
        })
    return render(request, 'admin_dashboard/users.html', {'users': user_data})

@superuser_required
def user_detail_view(request, user_id):
    user = get_object_or_404(User, id=user_id)
    sub = Subscription.objects.filter(user=user, status='Active').first()
    payments = Payment.objects.filter(user=user).order_by('-created_at')
    qrs = DynamicQRCode.objects.filter(user=user).exclude(qr_type='custom-url').order_by('-created_at')
    shorts = DynamicQRCode.objects.filter(user=user, qr_type='custom-url').order_by('-created_at')
    activity = ActivityLog.objects.filter(user=user).order_by('-created_at')
    
    context = {
        'target_user': user,
        'subscription': sub,
        'payments': payments,
        'qrs': qrs,
        'shorts': shorts,
        'activity': activity,
    }
    return render(request, 'admin_dashboard/user_detail.html', context)

@superuser_required
def subscriptions_view(request):
    subs = Subscription.objects.all().order_by('-created_at')
    return render(request, 'admin_dashboard/subscriptions.html', {'subscriptions': subs})

@superuser_required
def payments_view(request):
    payments = Payment.objects.all().order_by('-created_at')
    return render(request, 'admin_dashboard/payments.html', {'payments': payments})

@superuser_required
def plans_view(request):
    plans = Plan.objects.all()
    return render(request, 'admin_dashboard/plans.html', {'plans': plans})

@superuser_required
def qrcodes_view(request):
    qrs = DynamicQRCode.objects.exclude(qr_type='custom-url').order_by('-created_at')
    return render(request, 'admin_dashboard/qrcodes.html', {'qrcodes': qrs})

@superuser_required
def shorturls_view(request):
    shorts = DynamicQRCode.objects.filter(qr_type='custom-url').order_by('-created_at')
    return render(request, 'admin_dashboard/shorturls.html', {'shorturls': shorts})

@superuser_required
def reports_view(request):
    # Dummy data for charts, can be expanded later
    return render(request, 'admin_dashboard/reports.html', {})

@superuser_required
def activity_view(request):
    logs = ActivityLog.objects.all().order_by('-created_at')
    return render(request, 'admin_dashboard/activity.html', {'logs': logs})

@superuser_required
def settings_view(request):
    return render(request, 'admin_dashboard/settings.html', {})

def logout_view(request):
    logout(request)
    return redirect('converter:home')
