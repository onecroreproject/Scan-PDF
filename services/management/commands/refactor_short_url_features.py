from django.core.management.base import BaseCommand
from services.models import Plan, PlanSection, PlanSectionFeature, PlanFeature

class Command(BaseCommand):
    help = 'Refactors the Short URL features in the database (removes Bulk Import/Sticky, adds Header)'

    def handle(self, *args, **kwargs):
        self.stdout.write("Removing deprecated features...")
        
        # Remove old features
        PlanSectionFeature.objects.filter(name__icontains="Bulk Import").delete()
        PlanSectionFeature.objects.filter(name__icontains="Sticky").delete()
        self.stdout.write(self.style.SUCCESS("Successfully removed Bulk Import and Sticky features."))

        self.stdout.write("Updating Plan Features entitlements...")
        plans = Plan.objects.all()
        for plan in plans:
            # Update PlanSectionFeature
            sections = PlanSection.objects.filter(plan=plan, name__icontains="SHORT URL")
            if not sections.exists():
                continue
                
            short_url_section = sections.first()
            
            ps_feature, created = PlanSectionFeature.objects.get_or_create(
                section=short_url_section,
                name__icontains="Header",
                defaults={
                    "name": "Custom Headers",
                    "feature_type": "LIMIT",
                    "description": "Add a custom header to your short link.",
                    "display_order": 90,
                    "is_enabled": True
                }
            )
            
            ps_feature.name = "Custom Headers"
            ps_feature.feature_type = "LIMIT"
            
            # Find the globally defined Feature for PlanFeature
            from services.models import Feature
            global_feature, _ = Feature.objects.get_or_create(
                key="custom_header",
                defaults={
                    "name": "Custom Header",
                    "type": "LIMIT",
                    "is_public": False
                }
            )
            
            if plan.code == 'free':
                val = '5'
                ps_feature.monthly_value = 5
                ps_feature.yearly_value = 5
                ps_feature.is_unlimited = False
            elif plan.code == 'pro':
                val = '50'
                ps_feature.monthly_value = 50
                ps_feature.yearly_value = 50
                ps_feature.is_unlimited = False
            else:
                val = '-1'
                ps_feature.is_unlimited = True
                
            ps_feature.save()
            
            # Update PlanFeature
            pf, pf_created = PlanFeature.objects.get_or_create(
                plan=plan,
                feature=global_feature,
                defaults={
                    "value_text": val
                }
            )
            if not pf_created:
                pf.value_text = val
                pf.save()
                
            self.stdout.write(f"  - {plan.name} -> Header: {val}")

        self.stdout.write(self.style.SUCCESS("Database refactoring complete!"))
