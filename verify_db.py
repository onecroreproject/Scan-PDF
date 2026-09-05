import os
import django
from django.db.models import Count

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'core.settings')
django.setup()

from services.models import Plan, PlanFeature, FEATURE_CODES

print("=" * 50)
print("1. CHECKING FOR DUPLICATE PLANS")
print("=" * 50)
duplicate_codes = Plan.objects.values('code').annotate(count=Count('id')).filter(count__gt=1)
if duplicate_codes.exists():
    print(f"WARNING: Duplicates found for codes: {[d['code'] for d in duplicate_codes]}")
else:
    print("SUCCESS: No duplicate plans found.")
print()

print("=" * 50)
print("2. VERIFYING ACTIVE PLAN ROWS")
print("=" * 50)
plans = Plan.objects.filter(is_active=True).order_by('display_order')
for plan in plans:
    if plan.pricing_type == 'contact':
        print(f"Plan: {plan.name} (Code: {plan.code}, ID: {plan.id}) | Pricing: CONTACT US")
    else:
        print(f"Plan: {plan.name} (Code: {plan.code}, ID: {plan.id}) | Monthly: ₹{plan.monthly_price} | Yearly: ₹{plan.yearly_price}")
print()

print("=" * 50)
print("3. VERIFYING PLAN FEATURE ROWS")
print("=" * 50)
for plan in plans:
    print(f"--- {plan.name} FEATURES ---")
    pfs = PlanFeature.objects.filter(plan=plan, feature__key__in=FEATURE_CODES).select_related('feature')
    
    if not pfs.exists():
        print("  (No features found)")
        continue
        
    for pf in pfs:
        feat = pf.feature
        if pf.is_unlimited:
            print(f"  {feat.name}: Unlimited")
        elif feat.key == 'analytics':
            print(f"  {feat.name}: {pf.history_days} days")
        else:
            print(f"  {feat.name}: {pf.monthly_limit}/mo · {pf.yearly_limit}/yr")
    print()

print("=" * 50)
print("VERIFICATION COMPLETE")
print("=" * 50)
