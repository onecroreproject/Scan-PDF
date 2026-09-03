import json
from django.shortcuts import render, get_object_or_404, redirect
from django.contrib.auth import logout
from django.contrib.auth.models import User
from django.db.models import Sum, Count
from django.utils import timezone
from django.http import JsonResponse
from django.views.decorators.http import require_POST
from django.views.decorators.csrf import csrf_exempt
from datetime import timedelta
from services.models import (
    Subscription, Payment, Plan, ActivityLog,
    PlanSection, PlanSectionFeature
)
from dynamic_qr.models import DynamicQRCode
from .decorators import superuser_required


# ─────────────────────────────────────────────────────────────────────────────
# Dashboard & Standard Views
# ─────────────────────────────────────────────────────────────────────────────

@superuser_required
def dashboard_view(request):
    today = timezone.now().date()
    thirty_days_ago = timezone.now() - timedelta(days=30)

    context = {
        'total_users': User.objects.filter(is_superuser=False, is_staff=False).count(),
        'today_users': User.objects.filter(is_superuser=False, is_staff=False, date_joined__date=today).count(),
        'active_users': User.objects.filter(is_superuser=False, is_staff=False, is_active=True).count(),
        'inactive_users': User.objects.filter(is_superuser=False, is_staff=False, is_active=False).count(),

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
    users = User.objects.filter(is_superuser=False, is_staff=False).order_by('-date_joined')
    user_data = []
    for u in users:
        sub = Subscription.objects.filter(user=u, status='Active').first()
        qrs = DynamicQRCode.objects.filter(user=u).exclude(qr_type='custom-url').count()
        shorts = DynamicQRCode.objects.filter(user=u, qr_type='custom-url').count()
        user_data.append({'user': u, 'sub': sub, 'qr_count': qrs, 'short_count': shorts})
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
        'target_user': user, 'subscription': sub, 'payments': payments,
        'qrs': qrs, 'shorts': shorts, 'activity': activity,
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


# ─────────────────────────────────────────────────────────────────────────────
# Plans & Pricing
# ─────────────────────────────────────────────────────────────────────────────

@superuser_required
def plans_view(request):
    # Ensure old BUSINESS plan stays deactivated
    Plan.objects.filter(code='business').update(is_active=False)
    plans = Plan.objects.filter(is_active=True).prefetch_related(
        'sections__features'
    ).order_by('display_order', 'id')
    return render(request, 'admin_dashboard/plans.html', {'plans': plans})


@superuser_required
def plan_edit_view(request, plan_id):
    plan = get_object_or_404(Plan, id=plan_id)

    if request.method == 'POST':
        # Handle basic plan save (non-AJAX form submit)
        plan.name = request.POST.get('name', plan.name).strip()
        plan.description = request.POST.get('description', '').strip()
        plan.pricing_type = request.POST.get('pricing_type', 'fixed')
        plan.is_active = request.POST.get('is_active') == 'on'
        plan.is_popular = request.POST.get('is_popular') == 'on'
        try:
            plan.monthly_price = int(request.POST.get('monthly_price', 0))
        except (ValueError, TypeError):
            plan.monthly_price = 0
        try:
            plan.yearly_price = int(request.POST.get('yearly_price', 0))
        except (ValueError, TypeError):
            plan.yearly_price = 0
        plan.save()
        ActivityLog.objects.create(user=request.user, action=f"Updated Plan: {plan.name}")
        return redirect(f"{request.path}?saved=1")

    sections = plan.sections.prefetch_related('features').order_by('display_order', 'id')
    saved = request.GET.get('saved') == '1'

    # Build a fully JSON-safe data structure for the JavaScript Plan Builder.
    # No ORM objects, no QuerySets, no Decimals, no datetimes — only primitives.
    sections_json_data = _serialize_sections(sections)

    return render(request, 'admin_dashboard/plan_edit.html', {
        'plan': plan,
        'sections': sections,           # ORM queryset used for server-side HTML rendering
        'sections_json_data': sections_json_data,  # plain dicts used by JavaScript
        'saved': saved,
    })


def _serialize_sections(sections_qs):
    """
    Converts a PlanSection queryset (with prefetched features) into a plain
    list of dicts that is safe to pass through Django's json_script filter.
    No ORM/model instances remain in the output — only JSON-native types.
    """
    result = []
    for section in sections_qs:
        section_data = {
            'id': section.id,
            'name': section.name,
            'is_enabled': section.is_enabled,
            'display_order': section.display_order,
            'features': [],
        }
        for feat in section.features.all():
            section_data['features'].append({
                'id': feat.id,
                'name': feat.name,
                'feature_type': feat.feature_type,
                'is_unlimited': bool(feat.is_unlimited),
                'monthly_value': feat.monthly_value,   # int or None — both JSON-safe
                'yearly_value': feat.yearly_value,     # int or None — both JSON-safe
                'monthly_label': feat.monthly_label or '',
                'yearly_label': feat.yearly_label or '',
                'monthly_text': feat.monthly_text or '',
                'yearly_text': feat.yearly_text or '',
                'description': feat.description or '',
                'is_enabled': bool(feat.is_enabled),
                'display_order': feat.display_order,
                'section_id': section.id,
            })
        result.append(section_data)
    return result


# ─────────────────────────────────────────────────────────────────────────────
# Section AJAX CRUD
# ─────────────────────────────────────────────────────────────────────────────

@superuser_required
@require_POST
def section_create(request, plan_id):
    plan = get_object_or_404(Plan, id=plan_id)
    data = json.loads(request.body)
    name = data.get('name', '').strip()
    if not name:
        return JsonResponse({'error': 'Section name is required.'}, status=400)

    max_order = plan.sections.aggregate(m=Count('id'))['m'] or 0
    section = PlanSection.objects.create(
        plan=plan, name=name, display_order=max_order, is_enabled=True
    )
    ActivityLog.objects.create(user=request.user, action=f"Created section '{name}' in {plan.name}")
    return JsonResponse({
        'id': section.id, 'name': section.name,
        'display_order': section.display_order, 'is_enabled': section.is_enabled,
    })


@superuser_required
@require_POST
def section_update(request, section_id):
    section = get_object_or_404(PlanSection, id=section_id)
    data = json.loads(request.body)
    if 'name' in data:
        section.name = data['name'].strip() or section.name
    if 'is_enabled' in data:
        section.is_enabled = bool(data['is_enabled'])
    section.save()
    ActivityLog.objects.create(user=request.user, action=f"Updated section '{section.name}'")
    return JsonResponse({'success': True, 'name': section.name, 'is_enabled': section.is_enabled})


@superuser_required
@require_POST
def section_delete(request, section_id):
    section = get_object_or_404(PlanSection, id=section_id)
    plan_name = section.plan.name
    section_name = section.name
    section.delete()
    ActivityLog.objects.create(user=request.user, action=f"Deleted section '{section_name}' from {plan_name}")
    return JsonResponse({'success': True})


@superuser_required
@require_POST
def section_reorder(request, plan_id):
    plan = get_object_or_404(Plan, id=plan_id)
    data = json.loads(request.body)
    order = data.get('order', [])  # list of section IDs in new order
    for idx, sid in enumerate(order):
        PlanSection.objects.filter(id=sid, plan=plan).update(display_order=idx)
    return JsonResponse({'success': True})


# ─────────────────────────────────────────────────────────────────────────────
# Feature AJAX CRUD
# ─────────────────────────────────────────────────────────────────────────────

@superuser_required
@require_POST
def feature_create(request, section_id):
    section = get_object_or_404(PlanSection, id=section_id)
    data = json.loads(request.body)
    name = data.get('name', '').strip()
    if not name:
        return JsonResponse({'error': 'Feature name is required.'}, status=400)

    ftype = data.get('feature_type', 'BOOLEAN')
    is_unlimited = (ftype == 'UNLIMITED') or bool(data.get('is_unlimited'))

    def safe_int(val):
        try:
            return int(val)
        except (ValueError, TypeError):
            return None

    max_order = section.features.count()
    feat = PlanSectionFeature.objects.create(
        section=section,
        name=name,
        feature_type=ftype,
        is_unlimited=is_unlimited,
        monthly_value=safe_int(data.get('monthly_value')),
        yearly_value=safe_int(data.get('yearly_value')),
        monthly_label=data.get('monthly_label', '').strip(),
        yearly_label=data.get('yearly_label', '').strip(),
        monthly_text=data.get('monthly_text', '').strip(),
        yearly_text=data.get('yearly_text', '').strip(),
        description=data.get('description', '').strip(),
        display_order=max_order,
        is_enabled=True,
    )
    ActivityLog.objects.create(
        user=request.user,
        action=f"Created feature '{name}' in section '{section.name}' ({section.plan.name})"
    )
    return JsonResponse(_feature_to_dict(feat))


@superuser_required
@require_POST
def feature_update(request, feature_id):
    feat = get_object_or_404(PlanSectionFeature, id=feature_id)
    data = json.loads(request.body)

    def safe_int(val):
        try:
            return int(val)
        except (ValueError, TypeError):
            return None

    if 'name' in data:
        feat.name = data['name'].strip() or feat.name
    if 'feature_type' in data:
        feat.feature_type = data['feature_type']
        feat.is_unlimited = (feat.feature_type == 'UNLIMITED')
    if 'is_unlimited' in data:
        feat.is_unlimited = bool(data['is_unlimited'])
    if 'monthly_value' in data:
        feat.monthly_value = safe_int(data['monthly_value'])
    if 'yearly_value' in data:
        feat.yearly_value = safe_int(data['yearly_value'])
    if 'monthly_label' in data:
        feat.monthly_label = data['monthly_label'].strip()
    if 'yearly_label' in data:
        feat.yearly_label = data['yearly_label'].strip()
    if 'monthly_text' in data:
        feat.monthly_text = data['monthly_text'].strip()
    if 'yearly_text' in data:
        feat.yearly_text = data['yearly_text'].strip()
    if 'description' in data:
        feat.description = data['description'].strip()
    if 'is_enabled' in data:
        feat.is_enabled = bool(data['is_enabled'])
    feat.save()
    ActivityLog.objects.create(user=request.user, action=f"Updated feature '{feat.name}'")
    return JsonResponse(_feature_to_dict(feat))


@superuser_required
@require_POST
def feature_delete(request, feature_id):
    feat = get_object_or_404(PlanSectionFeature, id=feature_id)
    name = feat.name
    section_name = feat.section.name
    feat.delete()
    ActivityLog.objects.create(user=request.user, action=f"Deleted feature '{name}' from section '{section_name}'")
    return JsonResponse({'success': True})


@superuser_required
@require_POST
def feature_reorder(request, section_id):
    section = get_object_or_404(PlanSection, id=section_id)
    data = json.loads(request.body)
    order = data.get('order', [])
    for idx, fid in enumerate(order):
        PlanSectionFeature.objects.filter(id=fid, section=section).update(display_order=idx)
    return JsonResponse({'success': True})


def _feature_to_dict(feat):
    return {
        'id': feat.id,
        'name': feat.name,
        'feature_type': feat.feature_type,
        'is_unlimited': feat.is_unlimited,
        'monthly_value': feat.monthly_value,
        'yearly_value': feat.yearly_value,
        'monthly_label': feat.monthly_label,
        'yearly_label': feat.yearly_label,
        'monthly_text': feat.monthly_text,
        'yearly_text': feat.yearly_text,
        'description': feat.description,
        'is_enabled': feat.is_enabled,
        'display_order': feat.display_order,
        'monthly_display': feat.get_monthly_display(),
        'yearly_display': feat.get_yearly_display(),
    }


# ─────────────────────────────────────────────────────────────────────────────
# Other Admin Views
# ─────────────────────────────────────────────────────────────────────────────

@superuser_required
def features_view(request):
    """Legacy features view — redirect to plans since feature management is now inside each plan."""
    return redirect('custom_admin:plans')


@superuser_required
def qrcodes_view(request):
    qrs = DynamicQRCode.objects.exclude(qr_type='custom-url').order_by('-created_at')
    return render(request, 'admin_dashboard/qrcodes.html', {'qrcodes': qrs})


@superuser_required
def shorturls_view(request):
    """Admin short urls view removed. Redirect to dashboard."""
    return redirect('custom_admin:dashboard')


@superuser_required
def reports_view(request):
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


# ─────────────────────────────────────────────────────────────────────────────
# Hero Videos
# ─────────────────────────────────────────────────────────────────────────────
from converter.models import HeroVideo
from .forms import HeroVideoForm
from django.contrib import messages

@superuser_required
def hero_videos_view(request):
    videos = HeroVideo.objects.all()
    return render(request, 'admin_dashboard/hero_videos.html', {'videos': videos})

@superuser_required
def hero_video_add_view(request):
    if request.method == 'POST':
        form = HeroVideoForm(request.POST, request.FILES)
        if form.is_valid():
            form.save()
            messages.success(request, 'Hero video uploaded successfully.')
            return redirect('custom_admin:hero_videos')
    else:
        form = HeroVideoForm()
    return render(request, 'admin_dashboard/hero_video_form.html', {'form': form, 'title': 'Add Hero Video'})

@superuser_required
def hero_video_edit_view(request, video_id):
    video = get_object_or_404(HeroVideo, id=video_id)
    if request.method == 'POST':
        form = HeroVideoForm(request.POST, request.FILES, instance=video)
        if form.is_valid():
            form.save()
            messages.success(request, 'Hero video updated successfully.')
            return redirect('custom_admin:hero_videos')
    else:
        form = HeroVideoForm(instance=video)
    return render(request, 'admin_dashboard/hero_video_form.html', {'form': form, 'title': 'Edit Hero Video', 'video': video})

@superuser_required
def hero_video_delete_view(request, video_id):
    video = get_object_or_404(HeroVideo, id=video_id)
    if request.method == 'POST':
        video.delete()
        messages.success(request, 'Hero video deleted successfully.')
        return redirect('custom_admin:hero_videos')
    return render(request, 'admin_dashboard/hero_video_confirm_delete.html', {'video': video})
