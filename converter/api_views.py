"""
REST API views for the QR Code Generator page.

Endpoint: POST /api/qr/generate/

Supports all content types and customisation options available on the existing
/convert/qrcode-generator/ page and returns a structured JSON response with a
base64-encoded image so that any HTTP client receives everything in one call.

The underlying QR-generation logic (generate_qr_code, save_uploaded_file)
is reused from converter.utils exactly as the existing web page does.
"""

import os
import base64
import json
import re

from django.http import JsonResponse
from django.views.decorators.csrf import csrf_exempt

from .utils import generate_qr_code, save_uploaded_file


# ---------------------------------------------------------------------------
# Allowed option values - kept in sync with qrcode_generator.html template
# and the generate_qr_code() utility in converter/utils.py
# ---------------------------------------------------------------------------

VALID_BODY_STYLES = {
    "square", "rounded", "circle", "diamond", "dot", "small-square",
    "hline", "vline", "star", "cross", "leaf", "clover", "hexagon", "octagon",
}

VALID_EYE_STYLES = {
    "square", "rounded", "circle", "diamond", "leaf", "dot",
    "hexagon", "octagon", "star", "small-square",
}

VALID_BALL_STYLES = {
    "square", "rounded", "circle", "diamond", "star", "cross",
    "leaf", "hline", "vline", "clover",
}

VALID_GRADIENTS = {"none", "horizontal", "radial"}

VALID_OUTPUT_FORMATS = {"png", "jpg", "jpeg", "svg"}

VALID_CONTENT_TYPES = {
    "url", "text", "email", "phone", "sms", "vcard", "wifi", "location",
}

VALID_WIFI_ENCRYPTION = {"WPA", "WEP", "nopass"}

HEX_COLOR_RE = re.compile(r'^#([0-9A-Fa-f]{3}|[0-9A-Fa-f]{6})$')
MAX_LOGO_BYTES = 2 * 1024 * 1024  # 2 MB


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _success(data, message="QR code generated successfully"):
    """Return a structured success JSON response."""
    return JsonResponse({"success": True, "message": message, "data": data})


def _error(message, errors=None, status=400):
    """Return a structured error JSON response."""
    body = {"success": False, "message": message}
    if errors:
        body["errors"] = errors
    return JsonResponse(body, status=status)


def _validate_hex_color(value, field_name):
    """Return None if colour is valid hex, else an error string."""
    if not HEX_COLOR_RE.match(value):
        return f"{field_name} must be a valid hex colour (e.g. #000000)."
    return None


def _build_qr_text(content_type, payload):
    """
    Convert the structured content payload into the raw string encoded in the
    QR code.  Mirrors the buildQRData() JavaScript function in the template.

    Returns (qr_text, errors_dict).
    """
    errors = {}

    # -- URL -------------------------------------------------------------------
    if content_type == "url":
        url = payload.get("url", "").strip()
        if not url:
            errors["url"] = "url is required when content_type is 'url'."
        return url, errors

    # -- Plain text ------------------------------------------------------------
    if content_type == "text":
        text = payload.get("text", "").strip()
        if not text:
            errors["text"] = "text is required when content_type is 'text'."
        return text, errors

    # -- Email -----------------------------------------------------------------
    if content_type == "email":
        to = payload.get("email_to", "").strip()
        if not to:
            errors["email_to"] = "email_to is required when content_type is 'email'."
        subject = payload.get("email_subject", "").strip()
        body_text = payload.get("email_body", "").strip()
        qr_text = "mailto:{}".format(to)
        params = []
        if subject:
            from urllib.parse import quote
            params.append("subject={}".format(quote(subject)))
        if body_text:
            from urllib.parse import quote
            params.append("body={}".format(quote(body_text)))
        if params:
            qr_text += "?" + "&".join(params)
        return qr_text, errors

    # -- Phone -----------------------------------------------------------------
    if content_type == "phone":
        phone = payload.get("phone", "").strip()
        if not phone:
            errors["phone"] = "phone is required when content_type is 'phone'."
        return "tel:{}".format(phone), errors

    # -- SMS -------------------------------------------------------------------
    if content_type == "sms":
        number = payload.get("sms_number", "").strip()
        message = payload.get("sms_message", "").strip()
        if not number:
            errors["sms_number"] = "sms_number is required when content_type is 'sms'."
        return "smsto:{}:{}".format(number, message), errors

    # -- vCard -----------------------------------------------------------------
    if content_type == "vcard":
        fn = payload.get("first_name", "").strip()
        ln = payload.get("last_name", "").strip()
        if not fn and not ln:
            errors["first_name"] = (
                "At least first_name or last_name is required when content_type is 'vcard'."
            )

        def _v(key):
            return payload.get(key, "").strip()

        vcard = "BEGIN:VCARD\nVERSION:3.0\n"
        vcard += "N:{};{};;;\n".format(ln, fn)
        vcard += "FN:{} {}\n".format(fn, ln)
        if _v("organization"):
            vcard += "ORG:{}\n".format(_v("organization"))
        if _v("position"):
            vcard += "TITLE:{}\n".format(_v("position"))
        if _v("phone_work"):
            vcard += "TEL;TYPE=WORK:{}\n".format(_v("phone_work"))
        if _v("phone_mobile"):
            vcard += "TEL;TYPE=CELL:{}\n".format(_v("phone_mobile"))
        if _v("fax"):
            vcard += "TEL;TYPE=FAX:{}\n".format(_v("fax"))
        if _v("email"):
            vcard += "EMAIL:{}\n".format(_v("email"))
        if _v("website"):
            vcard += "URL:{}\n".format(_v("website"))
        adr_parts = ";".join([
            _v("street"), _v("city"), _v("state"), _v("zip_code"), _v("country")
        ])
        if adr_parts.replace(";", ""):
            vcard += "ADR;TYPE=WORK:;;{}\n".format(adr_parts)
        vcard += "END:VCARD"
        return vcard, errors

    # -- WiFi ------------------------------------------------------------------
    if content_type == "wifi":
        ssid = payload.get("wifi_ssid", "").strip()
        password = payload.get("wifi_password", "").strip()
        encryption = payload.get("wifi_encryption", "WPA").strip()
        hidden_raw = payload.get("wifi_hidden", False)
        hidden = hidden_raw in (True, "true", "1", "yes")

        if not ssid:
            errors["wifi_ssid"] = "wifi_ssid is required when content_type is 'wifi'."
        if encryption not in VALID_WIFI_ENCRYPTION:
            errors["wifi_encryption"] = (
                "wifi_encryption must be one of: {}.".format(
                    ", ".join(sorted(VALID_WIFI_ENCRYPTION))
                )
            )

        hidden_part = "H:true;" if hidden else ""
        qr_text = "WIFI:S:{};T:{};P:{};{};".format(ssid, encryption, password, hidden_part)
        return qr_text, errors

    # -- Location --------------------------------------------------------------
    if content_type == "location":
        lat = str(payload.get("latitude", "")).strip()
        lng = str(payload.get("longitude", "")).strip()
        if not lat or not lng:
            errors["latitude"] = (
                "Both latitude and longitude are required when content_type is 'location'."
            )
        return "geo:{},{}".format(lat, lng), errors

    # -- Fallback (should not reach here due to earlier validation) ------------
    errors["content_type"] = "Unsupported content_type: {!r}.".format(content_type)
    return "", errors


# ---------------------------------------------------------------------------
# Main REST API view
# ---------------------------------------------------------------------------

@csrf_exempt
def qr_generate_api(request):
    """
    GET  /api/v1/qr/generate/  -> 200 API information (no QR generated)
    POST /api/v1/qr/generate/  -> 200 JSON with base64-encoded QR image

    Legacy alias: /api/qr/generate/  (same view, both methods)
    """

    # ── GET: return endpoint/API metadata only, never generate a QR ───────
    if request.method == "GET":
        return JsonResponse(
            {
                "success": True,
                "message": "QR Code Generator API is available.",
                "data": {
                    "api_version": "v1",
                    "endpoint": "/api/v1/qr/generate/",
                    "allowed_methods": ["GET", "POST"],
                    "content_types": sorted(VALID_CONTENT_TYPES),
                    "output_formats": ["png", "jpg", "jpeg", "svg"],
                    "body_styles": sorted(VALID_BODY_STYLES),
                    "eye_styles": sorted(VALID_EYE_STYLES),
                    "ball_styles": sorted(VALID_BALL_STYLES),
                    "gradients": sorted(VALID_GRADIENTS),
                },
            },
            status=200,
        )

    # ── Reject everything that is not GET or POST ──────────────────────────
    if request.method != "POST":
        return JsonResponse(
            {
                "success": False,
                "message": "Method not allowed. Use GET or POST.",
            },
            status=405,
        )

    # ── POST: generate the QR code ─────────────────────────────────────────
    # ── 1. Parse request body ──────────────────────────────────────────────
    ct_header = request.content_type or ""

    if "application/json" in ct_header:
        try:
            payload = json.loads(request.body)
        except (ValueError, TypeError):
            return _error("Request body is not valid JSON.", status=400)
        logo_file = None
    else:
        # multipart/form-data or application/x-www-form-urlencoded
        payload = request.POST.dict()
        logo_file = request.FILES.get("logo")

    # ── 2. Validate customisation parameters ──────────────────────────────
    validation_errors = {}

    content_type = payload.get("content_type", "url").strip().lower()
    if content_type not in VALID_CONTENT_TYPES:
        validation_errors["content_type"] = (
            "content_type must be one of: {}.".format(
                ", ".join(sorted(VALID_CONTENT_TYPES))
            )
        )

    fg_color = payload.get("fg_color", "#000000").strip()
    bg_color = payload.get("bg_color", "#ffffff").strip()

    err = _validate_hex_color(fg_color, "fg_color")
    if err:
        validation_errors["fg_color"] = err

    err = _validate_hex_color(bg_color, "bg_color")
    if err:
        validation_errors["bg_color"] = err

    gradient = payload.get("gradient", "none").strip().lower()
    if gradient not in VALID_GRADIENTS:
        validation_errors["gradient"] = (
            "gradient must be one of: {}.".format(", ".join(sorted(VALID_GRADIENTS)))
        )

    style = payload.get("style", "square").strip().lower()
    if style not in VALID_BODY_STYLES:
        validation_errors["style"] = (
            "style must be one of: {}.".format(", ".join(sorted(VALID_BODY_STYLES)))
        )

    eye_style = payload.get("eye_style", "square").strip().lower()
    if eye_style not in VALID_EYE_STYLES:
        validation_errors["eye_style"] = (
            "eye_style must be one of: {}.".format(", ".join(sorted(VALID_EYE_STYLES)))
        )

    ball_style = payload.get("ball_style", "square").strip().lower()
    if ball_style not in VALID_BALL_STYLES:
        validation_errors["ball_style"] = (
            "ball_style must be one of: {}.".format(", ".join(sorted(VALID_BALL_STYLES)))
        )

    output_format = payload.get("output_format", "png").strip().lower()
    if output_format not in VALID_OUTPUT_FORMATS:
        validation_errors["output_format"] = (
            "output_format must be one of: {}.".format(
                ", ".join(sorted(VALID_OUTPUT_FORMATS))
            )
        )

    # Validate uploaded logo
    logo_path = None
    if logo_file is not None:
        allowed_logo_types = {
            "image/png", "image/jpeg", "image/jpg", "image/gif", "image/webp",
        }
        if logo_file.content_type not in allowed_logo_types:
            validation_errors["logo"] = (
                "Logo must be a valid image file (PNG, JPG, GIF, or WebP)."
            )
        elif logo_file.size > MAX_LOGO_BYTES:
            validation_errors["logo"] = "Logo file size must not exceed 2 MB."

    # ── 3. Build the raw QR text string from content-type-specific fields ──
    if "content_type" not in validation_errors:
        qr_text, content_errors = _build_qr_text(content_type, payload)
        validation_errors.update(content_errors)
    else:
        qr_text = ""

    if validation_errors:
        return _error("Validation failed.", errors=validation_errors, status=400)

    if not qr_text:
        return _error(
            "Validation failed.",
            errors={
                "content": (
                    "The resulting QR content is empty. "
                    "Please supply valid data for the chosen content_type."
                )
            },
            status=400,
        )

    # ── 4. Save the uploaded logo to a temp file (if provided) ────────────
    try:
        if logo_file is not None:
            logo_path = save_uploaded_file(logo_file)
    except Exception as exc:
        return _error(
            "Failed to process the uploaded logo: {}.".format(exc), status=500
        )

    # ── 5. Generate the QR code using the existing utility function ────────
    output_path = None
    try:
        output_path = generate_qr_code(
            text=qr_text,
            fg_color=fg_color,
            bg_color=bg_color,
            style=style,
            gradient_type=gradient,
            eye_style=eye_style,
            ball_style=ball_style,
            logo_path=logo_path,
            output_format=output_format,
        )
    except Exception as exc:
        return _error("QR code generation failed: {}.".format(exc), status=500)
    finally:
        # Always clean up the temp logo file
        if logo_path and os.path.exists(logo_path):
            try:
                os.remove(logo_path)
            except OSError:
                pass

    # ── 6. Read the generated file and encode it as base64 ────────────────
    try:
        with open(output_path, "rb") as fh:
            raw_bytes = fh.read()
        image_b64 = base64.b64encode(raw_bytes).decode("utf-8")
    except Exception as exc:
        return _error(
            "Failed to read generated QR code file: {}.".format(exc), status=500
        )
    finally:
        # Clean up the temp output file now that we have its content
        if output_path and os.path.exists(output_path):
            try:
                os.remove(output_path)
            except OSError:
                pass

    # ── 7. Determine MIME type and return structured JSON response ─────────
    fmt = output_format.lower()
    if fmt in ("jpg", "jpeg"):
        mime_type = "image/jpeg"
        file_extension = "jpg"
    elif fmt == "svg":
        mime_type = "image/svg+xml"
        file_extension = "svg"
    else:
        mime_type = "image/png"
        file_extension = "png"

    return _success(
        data={
            "image_base64": image_b64,
            "mime_type": mime_type,
            "file_extension": file_extension,
            "output_format": file_extension,
            "content_type": content_type,
            "qr_text": qr_text,
            "customisation": {
                "fg_color": fg_color,
                "bg_color": bg_color,
                "gradient": gradient,
                "style": style,
                "eye_style": eye_style,
                "ball_style": ball_style,
            },
        }
    )
