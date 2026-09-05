from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.conf import settings
from django.utils import timezone
from django.http import HttpResponseForbidden, Http404
from .models import Plan, Subscription, Payment
import random
import datetime


def _build_plan_features_for_display(plans):
    """
    Builds a display-ready features dict keyed by plan id for the pricing template.
    Returns: { plan_id: [{'name': ..., 'monthly_display': ..., 'yearly_display': ...}, ...] }

    Only shows the 9 active PlanFeature entries that are enabled.
    Display format: unlimited = 'Unlimited Feature', limit = 'N uses/mo', analytics = 'N days history'
    """
    from .models import PlanFeature, FEATURE_CODES

    plan_ids = [p.id for p in plans]
    pf_qs = PlanFeature.objects.filter(
        plan_id__in=plan_ids,
        enabled=True,
        feature__key__in=FEATURE_CODES,
        feature__is_active=True,
    ).select_related('feature').order_by('feature__display_order')

    result = {}
    for pf in pf_qs:
        pid = pf.plan_id
        if pid not in result:
            result[pid] = []

        feat = pf.feature

        if pf.is_unlimited:
            monthly_display = f"Unlimited {feat.name}"
            yearly_display = monthly_display
        elif feat.key == 'analytics':
            days = pf.history_days or 0
            monthly_display = f"{days} Days Analytics History"
            yearly_display = monthly_display
        elif pf.monthly_limit is not None or pf.yearly_limit is not None:
            m = pf.monthly_limit if pf.monthly_limit is not None else 0
            y = pf.yearly_limit if pf.yearly_limit is not None else 0
            monthly_display = f"{m} {feat.name} / month"
            yearly_display = f"{y} {feat.name} / year"
        else:
            monthly_display = feat.name
            yearly_display = feat.name

        result[pid].append({
            'name': feat.name,
            'monthly_display': monthly_display,
            'yearly_display': yearly_display,
        })

    return result


def pricing_view(request):
    """Renders the premium SaaS pricing page with active subscription info."""
    is_logged_in = request.user.is_authenticated and request.session.get('is_dqr_user')

    current_subscription = None
    if is_logged_in:
        from .plan_features import _get_active_subscription
        current_subscription = _get_active_subscription(request.user)

    plans = Plan.objects.filter(is_active=True).order_by('display_order', 'id')

    # Calculate max saving percentage for the badge
    max_saving = 0
    for plan in plans:
        if plan.pricing_type == 'fixed' and plan.monthly_price > 0:
            annualized = plan.monthly_price * 12
            if plan.yearly_price < annualized:
                saving = ((annualized - plan.yearly_price) / annualized) * 100
                if saving > max_saving:
                    max_saving = saving
    max_saving_percent = round(max_saving)

    # Build live feature display data from PlanFeature (not PlanSectionFeature)
    plan_features_display = _build_plan_features_for_display(plans)

    # Attach features list to each plan for template iteration
    plans_with_features = []
    for plan in plans:
        plans_with_features.append({
            'plan': plan,
            'features': plan_features_display.get(plan.id, []),
        })

    return render(request, 'services/pricing.html', {
        'is_logged_in': is_logged_in,
        'current_subscription': current_subscription,
        'plans': plans,
        'plans_with_features': plans_with_features,
        'max_saving_percent': max_saving_percent,
        'page_title': 'Pricing Plans — ScanPDF Services'
    })

@login_required(login_url='dynamic_qr:login')
def payment_confirm_view(request, plan_code, cycle):
    """Summarizes checkout and initiates payment simulation."""
    if not request.session.get('is_dqr_user'):
        return redirect('dynamic_qr:login')
        
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
