from .decorators import get_user_subscription

class SubscriptionMiddleware:
    """Middleware to attach active subscription context to the request object."""
    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        if request.user.is_authenticated and request.session.get('is_dqr_user'):
            request.subscription = get_user_subscription(request.user)
        else:
            request.subscription = None
            
        response = self.get_response(request)
        return response
