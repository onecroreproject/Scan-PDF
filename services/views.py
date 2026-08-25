from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.conf import settings
from django.utils import timezone
from django.http import HttpResponseForbidden, Http404
from .models import Plan, Subscription, Payment
import random
import datetime

def get_or_create_plans():
    """Initializes standard subscription plans if they do not exist."""
    plans_data = [
        {
            'name': 'FREE', 'code': 'free', 'monthly_price': 0, 'yearly_price': 0,
            'max_dynamic_qrs': 5, 'max_short_urls': 10, 'max_team_members': 0, 'max_domains': 0, 'max_api_requests': 0,
            'analytics_access': False, 'csv_export': False, 'gps_tracking': False, 'bulk_qr': False, 'webhooks': False
        },
        {
            'name': 'PRO', 'code': 'pro', 'monthly_price': 499, 'yearly_price': 4999,
            'max_dynamic_qrs': 100, 'max_short_urls': 100, 'max_team_members': 0, 'max_domains': 1, 'max_api_requests': 0,
            'analytics_access': True, 'csv_export': False, 'gps_tracking': False, 'bulk_qr': False, 'webhooks': False
        },
        {
            'name': 'BUSINESS', 'code': 'business', 'monthly_price': 999, 'yearly_price': 9999,
            'max_dynamic_qrs': 1000, 'max_short_urls': 1000, 'max_team_members': 3, 'max_domains': 3, 'max_api_requests': 5000,
            'analytics_access': True, 'csv_export': True, 'gps_tracking': True, 'bulk_qr': True, 'webhooks': False
        },
        {
            'name': 'BUSINESS+', 'code': 'business_plus', 'monthly_price': 1999, 'yearly_price': 19999,
            'max_dynamic_qrs': 999999, 'max_short_urls': 999999, 'max_team_members': 99, 'max_domains': 99, 'max_api_requests': 999999,
            'analytics_access': True, 'csv_export': True, 'gps_tracking': True, 'bulk_qr': True, 'webhooks': True
        }
    ]
    for data in plans_data:
        Plan.objects.get_or_create(code=data['code'], defaults=data)

def pricing_view(request):
    """Renders the premium SaaS pricing page with active subscription info."""
    get_or_create_plans()
    is_logged_in = request.user.is_authenticated and request.session.get('is_dqr_user')
    
    current_subscription = None
    if is_logged_in:
        current_subscription = Subscription.objects.filter(user=request.user, status='Active').first()
        if not current_subscription:
            # Auto-assign Free plan if user has no subscription at all
            free_plan = Plan.objects.get(code='free')
            current_subscription = Subscription.objects.create(
                user=request.user,
                plan=free_plan,
                status='Active',
                billing_cycle='monthly',
                payment_status='Paid'
            )
            
    return render(request, 'services/pricing.html', {
        'is_logged_in': is_logged_in,
        'current_subscription': current_subscription,
        'page_title': 'Pricing Plans — ScanPDF Services'
    })

@login_required(login_url='dynamic_qr:login')
def payment_confirm_view(request, plan_code, cycle):
    """Summarizes checkout and initiates payment simulation."""
    if not request.session.get('is_dqr_user'):
        return redirect('dynamic_qr:login')
        
    get_or_create_plans()
    plan = get_object_or_404(Plan, code=plan_code)
    
    if plan.code == 'free':
        # Free plan can be activated instantly without simulator
        Subscription.objects.filter(user=request.user, status='Active').update(status='Expired')
        Subscription.objects.create(
            user=request.user,
            plan=plan,
            status='Active',
            billing_cycle='monthly',
            payment_status='Paid'
        )
        return redirect('services:payment_success_view')

    price = plan.yearly_price if cycle == 'yearly' else plan.monthly_price
    
    return render(request, 'services/payment_confirm.html', {
        'plan': plan,
        'cycle': cycle,
        'price': price,
        'debug_mode': settings.DEBUG,
        'page_title': 'Confirm Payment — ScanPDF'
    })

@login_required(login_url='dynamic_qr:login')
def payment_simulate_view(request):
    """Simulates payment gateway return (Success or Fail). Strictly for DEBUG mode."""
    if not settings.DEBUG:
        raise Http404("Simulator is only available during development (DEBUG=True).")

    if request.method == 'POST':
        status = request.POST.get('status', 'failed')
        plan_code = request.POST.get('plan_code')
        cycle = request.POST.get('cycle')
        
        plan = get_object_or_404(Plan, code=plan_code)
        
        if status == 'success':
            # Create active subscription (cancel/expire other active ones)
            Subscription.objects.filter(user=request.user, status='Active').update(status='Expired')
            
            # End date calculation
            start_date = timezone.now()
            if cycle == 'yearly':
                end_date = start_date + datetime.timedelta(days=365)
            else:
                end_date = start_date + datetime.timedelta(days=30)
                
            sub = Subscription.objects.create(
                user=request.user,
                plan=plan,
                status='Active',
                start_date=start_date,
                end_date=end_date,
                billing_cycle=cycle,
                payment_status='Paid'
            )
            
            # Generate payment record
            amount = plan.yearly_price if cycle == 'yearly' else plan.monthly_price
            txn_id = f"TXN-{timezone.now().strftime('%Y%m%d')}-{random.randint(100000, 999999)}"
            
            payment = Payment.objects.create(
                user=request.user,
                subscription=sub,
                amount=amount,
                transaction_id=txn_id,
                payment_status='Paid',
                gateway='Simulator',
                payment_mode='Credit Card',
                receipt_number=f"REC-{random.randint(1000, 9999)}"
            )
            
            # Store in session for success display
            request.session['last_payment_id'] = payment.id
            return redirect('services:payment_success_view')
        else:
            return redirect('services:payment_failed_view')
            
    return redirect('services:pricing')

@login_required(login_url='dynamic_qr:login')
def payment_success_view(request):
    """Renders the Payment Successful confirmation page."""
    payment_id = request.session.pop('last_payment_id', None)
    payment = None
    if payment_id:
        payment = Payment.objects.filter(id=payment_id, user=request.user).first()
        
    # Fallback to last successful payment if none in session
    if not payment:
        payment = Payment.objects.filter(user=request.user, payment_status='Paid').order_by('-created_at').first()
        
    return render(request, 'services/payment_success.html', {
        'payment': payment,
        'subscription': payment.subscription if payment else None,
        'page_title': 'Payment Successful — ScanPDF'
    })

@login_required(login_url='dynamic_qr:login')
def payment_failed_view(request):
    """Renders the Payment Failed page."""
    return render(request, 'services/payment_failed.html', {
        'page_title': 'Payment Failed — ScanPDF'
    })

@login_required(login_url='dynamic_qr:login')
def payment_history_view(request):
    """Displays payment transactions table for the user."""
    payments = Payment.objects.filter(user=request.user).order_by('-created_at')
    return render(request, 'services/payment_history.html', {
        'payments': payments,
        'page_title': 'Billing History — ScanPDF'
    })

def support_view(request):
    """Placeholder view for customer support."""
    return render(request, 'services/placeholder_page.html', {
        'title': 'Premium Support',
        'subtitle': 'Dedicated support and service level agreements for enterprise accounts.',
        'page_title': 'Support — ScanPDF Services'
    })

def contact_view(request):
    """Placeholder view for contacting sales/support."""
    return render(request, 'services/placeholder_page.html', {
        'title': 'Contact Us',
        'subtitle': 'Get in touch with our solutions architects and support staff.',
        'page_title': 'Contact Us — ScanPDF Services'
    })
