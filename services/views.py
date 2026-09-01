from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.conf import settings
from django.utils import timezone
from django.http import HttpResponseForbidden, Http404
from .models import Plan, Subscription, Payment
import random
import datetime

def get_or_create_plans():
    """Initializes standard subscription plans and seeds default section/feature content."""
    from .models import PlanSection, PlanSectionFeature

    # Deactivate old BUSINESS ₹999 plan
    Plan.objects.filter(code='business').update(is_active=False)

    plans_data = [
        {
            'name': 'FREE', 'code': 'free', 'monthly_price': 0, 'yearly_price': 0,
            'pricing_type': 'fixed', 'description': 'Perfect to get started', 'display_order': 1,
            'is_active': True,
        },
        {
            'name': 'PRO', 'code': 'pro', 'monthly_price': 499, 'yearly_price': 4999,
            'pricing_type': 'fixed', 'description': 'For serious creators', 'is_popular': True,
            'display_order': 2, 'is_active': True,
        },
        {
            'name': 'BUSINESS+', 'code': 'business_plus', 'monthly_price': 0, 'yearly_price': 0,
            'pricing_type': 'contact', 'description': 'Enterprise scale', 'display_order': 3,
            'is_active': True,
        },
    ]
    for data in plans_data:
        Plan.objects.update_or_create(code=data['code'], defaults=data)

    # Seed default content only if plan has no sections yet
    _seed_default_sections()


def _seed_default_sections():
    """Seeds default PlanSection and PlanSectionFeature records for each plan if not yet configured."""
    from .models import PlanSection, PlanSectionFeature

    default_content = {
        'free': {
            'name': 'FREE', 'sections': [
                {
                    'name': 'SHORT URL', 'order': 0, 'features': [
                        {'name': 'Short URLs', 'type': 'LIMIT', 'mv': 100, 'yv': 1200, 'ml': 'Short URLs / month', 'yl': 'Short URLs / year'},
                        {'name': 'Link Redirection', 'type': 'UNLIMITED'},
                        {'name': 'Link Expiry', 'type': 'BOOLEAN'},
                        {'name': 'Edit Short Links', 'type': 'BOOLEAN'},
                        {'name': 'Custom Alias', 'type': 'BOOLEAN'},
                        {'name': 'UTM Builder & Tracking', 'type': 'BOOLEAN'},
                        {'name': 'Basic Link Analytics', 'type': 'BOOLEAN'},
                        {'name': 'Analytics History', 'type': 'TEXT', 'mt': '30 Days', 'yt': '365 Days', 'ml': 'Analytics History', 'yl': 'Analytics History'},
                        {'name': 'Password Protection', 'type': 'BOOLEAN'},
                        {'name': 'Collections', 'type': 'LIMIT', 'mv': 5, 'yv': 5, 'ml': 'Collections', 'yl': 'Collections'},
                    ]
                }
            ]
        },
        'pro': {
            'name': 'PRO', 'sections': [
                {
                    'name': 'SHORT URL', 'order': 0, 'features': [
                        {'name': 'Short URLs', 'type': 'LIMIT', 'mv': 1500, 'yv': 15000, 'ml': 'Short URLs / month', 'yl': 'Short URLs / year'},
                        {'name': 'Link Redirection', 'type': 'UNLIMITED'},
                        {'name': 'Link Expiry', 'type': 'UNLIMITED'},
                        {'name': 'Short Link Editing', 'type': 'UNLIMITED'},
                        {'name': 'Custom Aliases', 'type': 'BOOLEAN'},
                        {'name': 'Custom Domains', 'type': 'TEXT', 'mt': 'Up to 5', 'yt': 'Up to 5', 'ml': 'Custom Domains', 'yl': 'Custom Domains'},
                        {'name': 'Bulk URL Import', 'type': 'TEXT', 'mt': 'Up to 500 links', 'yt': 'Up to 500 links', 'ml': 'Bulk URL Import', 'yl': 'Bulk URL Import'},
                        {'name': 'Advanced Link Analytics', 'type': 'BOOLEAN'},
                        {'name': 'GPS Tracking', 'type': 'BOOLEAN'},
                        {'name': 'AI Insights', 'type': 'BOOLEAN'},
                        {'name': 'UTM Builder & Tracking', 'type': 'BOOLEAN'},
                        {'name': 'Collections', 'type': 'LIMIT', 'mv': 100, 'yv': 1000, 'ml': 'Collections', 'yl': 'Collections'},
                    ]
                }
            ]
        },
        'business_plus': {
            'name': 'BUSINESS+', 'sections': [
                {
                    'name': 'SHORT URL', 'order': 0, 'features': [
                        {'name': 'Short URLs', 'type': 'UNLIMITED'},
                        {'name': 'Link Redirection', 'type': 'UNLIMITED'},
                        {'name': 'Link Expiry', 'type': 'UNLIMITED'},
                        {'name': 'Short Link Editing', 'type': 'UNLIMITED'},
                        {'name': 'Custom Aliases', 'type': 'UNLIMITED'},
                        {'name': 'Custom Domains', 'type': 'UNLIMITED'},
                        {'name': 'Bulk URL Import', 'type': 'UNLIMITED'},
                        {'name': 'Custom Branded Short URLs', 'type': 'BOOLEAN'},
                        {'name': 'Advanced Analytics', 'type': 'BOOLEAN'},
                        {'name': 'AI Insights', 'type': 'BOOLEAN'},
                        {'name': 'Collections', 'type': 'UNLIMITED'},
                        {'name': 'Custom Analytics Reports', 'type': 'BOOLEAN'},
                    ]
                }
            ]
        },
    }

    for plan_code, plan_data in default_content.items():
        try:
            plan = Plan.objects.get(code=plan_code)
        except Plan.DoesNotExist:
            continue

        # Only seed if this plan has zero sections
        if plan.sections.exists():
            continue

        for sec_data in plan_data['sections']:
            section = PlanSection.objects.create(
                plan=plan,
                name=sec_data['name'],
                display_order=sec_data['order'],
                is_enabled=True,
            )
            for order_idx, feat in enumerate(sec_data['features']):
                ftype = feat['type']
                PlanSectionFeature.objects.create(
                    section=section,
                    name=feat['name'],
                    feature_type=ftype,
                    is_unlimited=(ftype == 'UNLIMITED'),
                    monthly_value=feat.get('mv'),
                    yearly_value=feat.get('yv'),
                    monthly_label=feat.get('ml', ''),
                    yearly_label=feat.get('yl', ''),
                    monthly_text=feat.get('mt', ''),
                    yearly_text=feat.get('yt', ''),
                    display_order=order_idx,
                    is_enabled=True,
                )


def pricing_view(request):
    """Renders the premium SaaS pricing page with active subscription info."""
    get_or_create_plans()
    is_logged_in = request.user.is_authenticated

    current_subscription = None
    if is_logged_in:
        from .limit_service import PlanLimitService
        current_subscription = PlanLimitService.get_active_subscription(request.user)

    plans = Plan.objects.filter(is_active=True).prefetch_related(
        'sections__features'
    ).order_by('display_order', 'id')

    return render(request, 'services/pricing.html', {
        'is_logged_in': is_logged_in,
        'current_subscription': current_subscription,
        'plans': plans,
        'page_title': 'Pricing Plans — ScanPDF Services'
    })

@login_required(login_url='dynamic_qr:login')
def payment_confirm_view(request, plan_code, cycle):
    """Summarizes checkout and initiates payment simulation."""
    if not request.user.is_authenticated:
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
