from .models import ActivityLog

def log_activity(user, action, request=None, status='Success'):
    """Logs system and user actions to the database."""
    ip = None
    if request:
        x_forwarded_for = request.META.get('HTTP_X_FORWARDED_FOR')
        if x_forwarded_for:
            ip = x_forwarded_for.split(',')[0].strip()
        else:
            ip = request.META.get('REMOTE_ADDR')
            
    ActivityLog.objects.create(
        user=user if user and user.is_authenticated else None,
        action=action,
        ip_address=ip,
        status=status
    )
