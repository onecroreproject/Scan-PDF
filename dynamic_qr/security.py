import hashlib
import logging
from django.core.cache import cache, caches

try:
    sec_cache = caches['security']
except Exception:
    sec_cache = cache
from .utils import get_client_ip

logger = logging.getLogger(__name__)

def get_security_client_key(request, resource_id=None):
    """
    Returns a stable privacy-preserving digest based solely on the trusted client IP.
    Does NOT include User-Agent to prevent rate-limit bypass.
    """
    client_ip = get_client_ip(request)
    salt = "dqr_sec_key_v1"
    raw_key = f"{salt}:{client_ip}:{resource_id or 'global'}"
    return hashlib.sha256(raw_key.encode('utf-8')).hexdigest()[:32]

class PasswordAttemptLimiter:
    """
    Production-safe atomic brute-force limiter for passwords.
    """
    THRESHOLD = 5
    LOCKOUT_SECONDS = 300  # 5 minutes
    
    @classmethod
    def get_key(cls, qr_id, client_key):
        return f"shorturl:pwd_fail:{qr_id}:{client_key}"

    @classmethod
    def get_lock_key(cls, qr_id, client_key):
        return f"shorturl:pwd_lock:{qr_id}:{client_key}"
        
    @classmethod
    def is_locked(cls, qr_id, client_key):
        try:
            return bool(sec_cache.get(cls.get_lock_key(qr_id, client_key)))
        except Exception as e:
            # Fail-closed for security if cache is down
            logger.error(f"Password cache read failed: {e}")
            return True 
            
    @classmethod
    def record_failure(cls, qr_id, client_key):
        fail_key = cls.get_key(qr_id, client_key)
        lock_key = cls.get_lock_key(qr_id, client_key)
        
        try:
            # Atomic increment
            try:
                fails = sec_cache.incr(fail_key)
            except ValueError:
                # Key doesn't exist
                sec_cache.set(fail_key, 1, timeout=cls.LOCKOUT_SECONDS)
                fails = 1
                
            if fails >= cls.THRESHOLD:
                sec_cache.set(lock_key, '1', timeout=cls.LOCKOUT_SECONDS)
        except Exception as e:
            logger.error(f"Password cache write failed: {e}")

    @classmethod
    def clear_failures(cls, qr_id, client_key):
        try:
            sec_cache.delete_many([cls.get_key(qr_id, client_key), cls.get_lock_key(qr_id, client_key)])
        except Exception as e:
            logger.error(f"Password cache delete failed: {e}")

class RedirectAbuseLimiter:
    """
    Lightweight limiter for public endpoints to prevent enumeration and DB abuse.
    """
    from django.conf import settings
    BURST_LIMIT = getattr(settings, 'REDIRECT_RATE_LIMIT_BURST', 600)  # 600 requests
    WINDOW_SECONDS = getattr(settings, 'REDIRECT_RATE_LIMIT_WINDOW', 60) # per 1 minute
    
    @classmethod
    def is_rate_limited(cls, client_key):
        key = f"shorturl:redir_rate:{client_key}"
        try:
            try:
                count = cache.incr(key)
            except ValueError:
                cache.set(key, 1, timeout=cls.WINDOW_SECONDS)
                count = 1
                
            if count > cls.BURST_LIMIT:
                return True
        except Exception as e:
            # Fail-open for general redirects to preserve availability
            logger.error(f"Redirect cache failed (failing open): {e}")
        return False
