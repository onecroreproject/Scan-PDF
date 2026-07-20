from django.contrib.auth.decorators import user_passes_test
from django.shortcuts import redirect

def superuser_required(view_func):
    def check_perms(user):
        return user.is_active and user.is_superuser
    
    actual_decorator = user_passes_test(
        check_perms,
        login_url='dynamic_qr:login',
    )
    
    def wrapped_view(request, *args, **kwargs):
        if not request.user.is_authenticated:
            return redirect('dynamic_qr:login')
        if not (request.user.is_active and request.user.is_superuser):
            return redirect('dynamic_qr:dashboard')
        return actual_decorator(view_func)(request, *args, **kwargs)
        
    return wrapped_view
