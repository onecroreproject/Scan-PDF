# Generated manually for services app — adds PlanFeature entitlement fields and Feature.section/is_active

import django.utils.timezone
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('services', '0009_remove_plansectionfeature_feature_key'),
    ]

    operations = [
        # Add new fields to Feature model
        migrations.AddField(
            model_name='feature',
            name='section',
            field=models.CharField(default='SHORT URL', help_text='Section this feature belongs to', max_length=100),
        ),
        migrations.AddField(
            model_name='feature',
            name='is_active',
            field=models.BooleanField(default=True),
        ),
        migrations.AddField(
            model_name='feature',
            name='updated_at',
            field=models.DateTimeField(auto_now=True),
        ),

        # Add new entitlement fields to PlanFeature
        migrations.AddField(
            model_name='planfeature',
            name='enabled',
            field=models.BooleanField(default=False, help_text='Whether this feature is enabled for this plan'),
        ),
        migrations.AddField(
            model_name='planfeature',
            name='limit',
            field=models.IntegerField(blank=True, help_text='Monthly usage limit (null = no numeric limit)', null=True),
        ),
        migrations.AddField(
            model_name='planfeature',
            name='is_unlimited',
            field=models.BooleanField(default=False, help_text='If True, ignore limit entirely'),
        ),
        migrations.AddField(
            model_name='planfeature',
            name='limit_period',
            field=models.CharField(
                choices=[('monthly', 'Monthly'), ('none', 'No Period')],
                default='monthly',
                help_text='Period over which the limit applies',
                max_length=10,
            ),
        ),
        migrations.AddField(
            model_name='planfeature',
            name='history_days',
            field=models.IntegerField(
                blank=True,
                help_text='For analytics feature: how many days of history are accessible',
                null=True,
            ),
        ),
        migrations.AddField(
            model_name='planfeature',
            name='created_at',
            field=models.DateTimeField(auto_now_add=True, default=django.utils.timezone.now),
            preserve_default=False,
        ),
        migrations.AddField(
            model_name='planfeature',
            name='updated_at',
            field=models.DateTimeField(auto_now=True),
        ),
    ]
