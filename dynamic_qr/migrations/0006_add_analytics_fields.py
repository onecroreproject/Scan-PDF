# Generated manually to add Analytics Fields

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('dynamic_qr', '0005_remove_dynamicqrcode_header_data_and_more'),
    ]

    operations = [
        migrations.AddField(
            model_name='qranalytics',
            name='http_status',
            field=models.IntegerField(default=200),
        ),
        migrations.AddField(
            model_name='qranalytics',
            name='redirect_result',
            field=models.CharField(choices=[('redirect_success', 'Successful Redirect'), ('password_required', 'Password Challenge'), ('password_failed', 'Password Failure'), ('expired', 'Expired Link Attempt'), ('disabled', 'Disabled Link Attempt'), ('gps_required', 'GPS Permission Required'), ('gps_denied', 'GPS Permission Denied'), ('invalid_link', 'Invalid Link Request'), ('bot_request', 'Bot Request Ignored')], db_index=True, default='redirect_success', max_length=30),
        ),
    ]
