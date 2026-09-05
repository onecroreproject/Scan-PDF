from django.db import migrations, models


class Migration(migrations.Migration):
    dependencies = [
        ('dynamic_qr', '0006_add_analytics_fields'),
    ]

    operations = [
        migrations.AddField(
            model_name='qranalytics',
            name='location_source',
            field=models.CharField(default='unknown', max_length=20),
        ),
        migrations.AddField(
            model_name='qranalytics',
            name='gps_permission',
            field=models.CharField(default='not_required', max_length=20),
        ),
        migrations.AddField(
            model_name='qranalytics',
            name='gps_latitude',
            field=models.FloatField(blank=True, null=True),
        ),
        migrations.AddField(
            model_name='qranalytics',
            name='gps_longitude',
            field=models.FloatField(blank=True, null=True),
        ),
        migrations.AddField(
            model_name='qranalytics',
            name='gps_accuracy',
            field=models.FloatField(blank=True, null=True),
        ),
        migrations.AddField(
            model_name='qranalytics',
            name='gps_captured_at',
            field=models.DateTimeField(blank=True, null=True),
        ),
    ]