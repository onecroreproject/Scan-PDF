import logging
from django.conf import settings
from django.core.mail import EmailMultiAlternatives
from django.template.loader import render_to_string

logger = logging.getLogger(__name__)


def _enquiry_context(enquiry):
    source_label = 'Contact' if enquiry.source == 'CONTACT' else 'Help'
    return {
        'name': enquiry.name,
        'email': enquiry.email,
        'phone': enquiry.phone or 'Not provided',
        'category': enquiry.category,
        'subject': enquiry.subject,
        'message': enquiry.message,
        'ticket_id': enquiry.ticket_id,
        'source': source_label,
        'created_at': enquiry.created_at,
        'admin_email': getattr(settings, 'CONTACT_RECEIVER_EMAIL', None) or getattr(settings, 'DEFAULT_FROM_EMAIL', None) or settings.EMAIL_HOST_USER,
    }


def send_admin_enquiry_email(enquiry):
    receiver_email = getattr(settings, 'CONTACT_RECEIVER_EMAIL', None) or getattr(settings, 'DEFAULT_FROM_EMAIL', None) or settings.EMAIL_HOST_USER
    if not receiver_email:
        logger.warning('No contact receiver email configured for enquiry notification.')
        return False

    context = _enquiry_context(enquiry)
    subject = f"[ScanPDF {('Contact' if enquiry.source == 'CONTACT' else 'Help')}] {enquiry.category} - {enquiry.subject}"
    text_content = render_to_string('emails/contact_admin.txt', context)
    html_content = render_to_string('emails/contact_admin.html', context)

    from_email = getattr(settings, 'DEFAULT_FROM_EMAIL', None) or settings.EMAIL_HOST_USER
    msg = EmailMultiAlternatives(
        subject,
        text_content,
        from_email,
        [receiver_email],
        reply_to=[enquiry.email],
    )
    msg.attach_alternative(html_content, 'text/html')
    try:
        msg.send(fail_silently=True)
        return True
    except Exception:
        logger.exception('Failed to send admin enquiry email for ticket %s', enquiry.ticket_id)
        return False


def send_user_acknowledgement_email(enquiry):
    if not enquiry.email:
        logger.warning('No email address available to send acknowledgement for ticket %s', enquiry.ticket_id)
        return False

    context = _enquiry_context(enquiry)
    subject = 'We received your message | ScanPDF'
    text_content = render_to_string('emails/contact_confirmation.txt', context)
    html_content = render_to_string('emails/contact_confirmation.html', context)

    from_email = getattr(settings, 'DEFAULT_FROM_EMAIL', None) or settings.EMAIL_HOST_USER
    msg = EmailMultiAlternatives(
        subject,
        text_content,
        from_email,
        [enquiry.email],
    )
    msg.attach_alternative(html_content, 'text/html')
    try:
        msg.send(fail_silently=True)
        return True
    except Exception:
        logger.exception('Failed to send acknowledgement email for ticket %s', enquiry.ticket_id)
        return False
