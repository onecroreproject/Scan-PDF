import logging
import random
import datetime
import re

from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.conf import settings
from django.utils import timezone
from django.http import HttpResponseForbidden, Http404, JsonResponse
from django.core.cache import cache
from django.views.decorators.http import require_http_methods

from .models import Plan, Subscription, Payment, ContactEnquiry
from .forms import ContactEnquiryForm
from .email_service import send_admin_enquiry_email, send_user_acknowledgement_email

logger = logging.getLogger(__name__)


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

def support_legacy_redirect(request):
    return redirect('services:help', permanent=False)


def _get_client_ip(request):
    x_forwarded_for = request.META.get('HTTP_X_FORWARDED_FOR')
    if x_forwarded_for:
        return x_forwarded_for.split(',')[0].strip()
    return request.META.get('REMOTE_ADDR') or ''


def _make_ticket_id(source):
    prefix = 'SCN-HLP' if source == ContactEnquiry.SOURCE_HELP else 'SCN-CON'
    date_part = timezone.now().strftime('%Y%m%d')
    random_part = ''.join(random.choices('ABCDEFGHJKLMNPQRSTUVWXYZ23456789', k=5))
    return f'{prefix}-{date_part}-{random_part}'


def _rate_limited(request):
    ip_address = _get_client_ip(request)
    if not ip_address:
        return False
    cache_key = f'enquiry_rate_limit:{ip_address}'
    count = cache.get(cache_key, 0)
    if count >= 5:
        return True
    cache.set(cache_key, count + 1, timeout=600)
    return False


def _render_enquiry_form(request, template_name, form_class, source, title, subtitle, extra_context=None):
    form = form_class(request.POST or None, source=source)
    success = False
    ticket_id = None
    ajax_requested = request.headers.get('X-Requested-With') == 'XMLHttpRequest'

    if request.method == 'POST' and not _rate_limited(request):
        try:
            if form.is_valid():
                enquiry = form.save(commit=False)
                enquiry.source = source
                enquiry.ip_address = _get_client_ip(request)
                enquiry.user_agent = request.META.get('HTTP_USER_AGENT', '')[:500]
                enquiry.ticket_id = _make_ticket_id(source)
                while ContactEnquiry.objects.filter(ticket_id=enquiry.ticket_id).exists():
                    enquiry.ticket_id = _make_ticket_id(source)
                enquiry.save()

                admin_sent = send_admin_enquiry_email(enquiry)
                user_sent = send_user_acknowledgement_email(enquiry)
                if not admin_sent:
                    logger.warning('Admin enquiry email failed for ticket %s', enquiry.ticket_id)
                if not user_sent:
                    logger.warning('User acknowledgement email failed for ticket %s', enquiry.ticket_id)

                success = True
                ticket_id = enquiry.ticket_id
                form = form_class(source=source)

                if ajax_requested:
                    if not admin_sent or not user_sent:
                        return JsonResponse({
                            'success': True,
                            'status': 'saved_email_failed',
                            'message': 'Request received. Confirmation email may be delayed.',
                            'ticket_id': ticket_id,
                            'source': source,
                            'submitted_email': enquiry.email,
                        })
                    return JsonResponse({
                        'success': True,
                        'status': 'sent',
                        'message': 'Message sent successfully.',
                        'ticket_id': ticket_id,
                        'source': source,
                        'submitted_email': enquiry.email,
                    })
            else:
                if ajax_requested:
                    errors = {}
                    for field, field_errors in form.errors.items():
                        if field == '__all__':
                            errors['non_field_errors'] = list(field_errors)
                        else:
                            errors[field] = list(field_errors)
                    return JsonResponse({
                        'success': False,
                        'status': 'validation_error',
                        'errors': errors,
                        'message': 'Please check your details and try again.',
                    }, status=400)
        except Exception:
            logger.exception('Unexpected error while processing enquiry for source %s', source)
            if ajax_requested:
                return JsonResponse({
                    'success': False,
                    'status': 'server_error',
                    'message': 'Something went wrong. Please try again.',
                }, status=500)
            raise
    elif request.method == 'POST' and _rate_limited(request):
        if ajax_requested:
            return JsonResponse({
                'success': False,
                'status': 'rate_limited',
                'message': 'Too many attempts. Please wait a moment and try again.',
            }, status=429)

    context = {
        'title': title,
        'subtitle': subtitle,
        'page_title': f'{title} — ScanPDF Services',
        'form': form,
        'success': success,
        'ticket_id': ticket_id,
        'contact_email': getattr(settings, 'CONTACT_RECEIVER_EMAIL', None) or getattr(settings, 'DEFAULT_FROM_EMAIL', None) or settings.EMAIL_HOST_USER,
        'business_email': getattr(settings, 'BUSINESS_EMAIL', None) or getattr(settings, 'CONTACT_RECEIVER_EMAIL', None) or getattr(settings, 'DEFAULT_FROM_EMAIL', None) or '',
        'response_time': getattr(settings, 'SUPPORT_RESPONSE_TIME', 'We usually respond within 1–2 business days.'),
    }
    if extra_context:
        context.update(extra_context)
    return render(request, template_name, context)


def help_view(request):
    faq_items = [
        {'question': 'How do I use ScanPDF tools?', 'answer': 'Open any tool from the main navigation, upload your file, configure the settings, and click the action button to process it. Most tools work in a few guided steps.'},
        {'question': 'Do I need an account?', 'answer': 'Many quick tools can be used without an account, but premium features, saved settings, and analytics are available for registered users.'},
        {'question': 'Why is my file not uploading?', 'answer': 'Check the file size, format, and browser permissions. Most ScanPDF tools accept common PDF, image, and video formats. If the issue continues, use the Help form below.'},
        {'question': 'What file formats are supported?', 'answer': 'ScanPDF supports common PDF, JPG, PNG, WEBP, MP4, MOV, and other standard formats depending on the selected tool. The tool page usually lists supported formats.'},
        {'question': 'Is my uploaded file secure?', 'answer': 'We process files securely and do not keep them longer than needed for the task. Uploaded content is handled according to our privacy and security practices.'},
        {'question': 'How long are uploaded files stored?', 'answer': 'File retention depends on the tool and active user policy. Temporary files are typically cleaned shortly after processing, while premium account data may be retained longer.'},
        {'question': 'How do I create a short URL?', 'answer': 'Open the Short URL tool, paste your destination link, choose optional settings like expiry or password protection, and generate a shareable short link.'},
        {'question': 'How do I generate a QR code?', 'answer': 'Go to the Dynamic QR tool, select the QR type, fill in the fields, and export or share the generated code.'},
        {'question': 'What is included in the Free plan?', 'answer': 'The Free plan includes basic access to public tools with limited usage and feature availability. You can review the current plan details in the pricing section.'},
        {'question': 'What is included in the Pro plan?', 'answer': 'The Pro plan adds higher limits, premium features, and a better processing experience for individuals and teams. See the pricing section for the latest details.'},
        {'question': 'How do I upgrade my account?', 'answer': 'Visit the pricing page and choose a plan that fits your needs. If you are already logged in, you can move through the checkout and billing flow directly.'},
        {'question': 'I paid but my account is not upgraded. What should I do?', 'answer': 'Please check the payment status and then contact the support team with your transaction details. Include the email used for the account and the reference number if available.'},
        {'question': 'How can I report a bug?', 'answer': 'Use the Help form below and select Bug Report as the category, or contact the support team with a clear description of the problem and the tool affected.'},
        {'question': 'How can I contact ScanPDF?', 'answer': 'You can use the Help form on this page or visit the Contact Us page for general inquiries, business requests, and support follow-ups.'},
    ]

    help_cards = [
        {'icon': 'file-text', 'title': 'PDF Tools', 'description': 'Convert, compress, merge, split, and edit PDF documents.'},
        {'icon': 'image', 'title': 'Image Tools', 'description': 'Resize, compress, convert, watermark, and optimize images.'},
        {'icon': 'video', 'title': 'Video Tools', 'description': 'Trim, merge, edit, and optimize video files with ease.'},
        {'icon': 'qr-code', 'title': 'QR Code', 'description': 'Create dynamic QR codes for links, contact details, and more.'},
        {'icon': 'link-2', 'title': 'Short URL', 'description': 'Generate short, branded, and trackable URLs for sharing.'},
        {'icon': 'user-circle', 'title': 'Account & Login', 'description': 'Manage your account, sign in, and access your profile settings.'},
        {'icon': 'badge-dollar-sign', 'title': 'Pricing & Plans', 'description': 'Compare plan features, upgrades, and usage limits.'},
        {'icon': 'receipt', 'title': 'Billing', 'description': 'Learn about payment status, invoices, and renewals.'},
        {'icon': 'shield-check', 'title': 'Privacy & Security', 'description': 'Review file handling, retention, and account protections.'},
        {'icon': 'wrench', 'title': 'Technical Issues', 'description': 'Troubleshoot browser, upload, and processing-related errors.'},
    ]

    form = ContactEnquiryForm(source=ContactEnquiry.SOURCE_HELP)
    return _render_enquiry_form(request, 'services/help.html', ContactEnquiryForm, ContactEnquiry.SOURCE_HELP, 'How Can We Help?', 'Find answers, troubleshoot issues, or contact our support team.', {
        'help_cards': help_cards,
        'faq_items': faq_items,
        'form': form,
    })


def contact_view(request):
    form = ContactEnquiryForm(source=ContactEnquiry.SOURCE_CONTACT)
    return _render_enquiry_form(request, 'services/contact.html', ContactEnquiryForm, ContactEnquiry.SOURCE_CONTACT, "Let's Talk", 'Have a question, feedback, business enquiry, or need technical assistance? Send us a message and our team will get back to you.', {
        'form': form,
        'contact_email': getattr(settings, 'CONTACT_RECEIVER_EMAIL', None) or getattr(settings, 'DEFAULT_FROM_EMAIL', None) or settings.EMAIL_HOST_USER,
        'business_email': getattr(settings, 'BUSINESS_EMAIL', None) or getattr(settings, 'CONTACT_RECEIVER_EMAIL', None) or getattr(settings, 'DEFAULT_FROM_EMAIL', None) or '',
        'response_time': getattr(settings, 'SUPPORT_RESPONSE_TIME', 'We usually respond within 1–2 business days.'),
    })
