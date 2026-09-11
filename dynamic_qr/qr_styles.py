"""Validated QR style registry shared by persistence and rendering."""

import re

DEFAULT_STYLE_CONFIG = {
    'version': 1,
    'patterns': {
        'body': 'square',
        'outer_eye': 'square',
        'inner_eye': 'square',
    },
    'colors': {
        'body': '#000000',
        'outer_eye': '#000000',
        'inner_eye': '#000000',
        'background': '#FFFFFF',
    },
    'frame': {
        'type': 'none',
        'text': '',
        'text_size': 18,
        'font': 'Arial',
        'use_custom_colors': False,
        'frame_color': '#000000',
        'text_color': '#FFFFFF',
    },
}

FRAME_TYPES = {
    'none', 'simple_circle', 'rounded_square', 'simple_square', 'bottom_bar',
    'top_bar', 'top_bottom_bar', 'label_bottom', 'label_top', 'banner_top',
    'banner_bottom', 'speech_top', 'speech_bottom', 'corners',
    'rounded_corners', 'scan_me_top', 'scan_me_bottom', 'ribbon_top',
    'ribbon_bottom', 'ticket', 'receipt', 'device_phone', 'device_tablet',
    'badge', 'card', 'poster',
}

FRAME_FONTS = {
    'Times New Roman', 'Georgia', 'Arial', 'Helvetica', 'Verdana', 'Tahoma',
    'Trebuchet MS', 'Courier New', 'Monaco', 'Comic Sans MS', 'Impact',
    'Baskerville', 'Papyrus', 'Lucida Sans', 'Gill Sans',
}

BODY_PATTERNS = {
    'square', 'rounded_square', 'circle', 'diamond', 'small_circle',
    'small_square', 'horizontal', 'vertical', 'star', 'plus', 'rounded',
    'dot', 'hline', 'vline', 'small-square',
    'connected', 'hexagon', 'octagon', 'cross', 'flower', 'clover',
    'sparkle', 'soft_diamond', 'burst', 'droplet', 'leaf', 'angled_square',
    'tilted_square', 'four_dots', 'pixel_plus', 'soft_cross', 'corner_round',
    'blob',
}

OUTER_EYE_STYLES = {
    'square', 'circle', 'rounded', 'rounded_square', 'diamond', 'soft_square', 'leaf',
    'dot', 'small-square',
    'hexagon', 'octagon', 'star', 'cut_corner', 'blob', 'squircle',
}

INNER_EYE_STYLES = {
    'square', 'circle', 'rounded', 'rounded_square', 'diamond', 'soft_square', 'leaf',
    'dot', 'small-square',
    'star', 'hexagon', 'octagon', 'clover', 'sparkle', 'blob',
}

_COLOR_RE = re.compile(r'^#[0-9a-fA-F]{6}$')


def _color(value, fallback):
    value = str(value or '').strip()
    return value.upper() if _COLOR_RE.fullmatch(value) else fallback


def normalize_frame_config(value=None, *, legacy_style='none', legacy_text='',
                           legacy_font='Arial', legacy_text_color='#FFFFFF'):
    """Validate frame settings while accepting the legacy frame keys."""
    source = value if isinstance(value, dict) else {}
    frame = source.get('frame') if isinstance(source.get('frame'), dict) else source
    frame_type = frame.get('type', frame.get('frame_style', legacy_style))
    if frame_type not in FRAME_TYPES:
        frame_type = 'none'
    text = str(frame.get('text', frame.get('frame_text', legacy_text)) or '')[:40]
    try:
        text_size = int(frame.get('text_size', 18))
    except (TypeError, ValueError):
        text_size = 18
    font = frame.get('font', frame.get('frame_font', legacy_font))
    if font not in FRAME_FONTS:
        font = 'Arial'
    return {
        'type': frame_type,
        'text': text,
        'text_size': max(10, min(40, text_size)),
        'font': font,
        'use_custom_colors': bool(frame.get('use_custom_colors', False)),
        'frame_color': _color(frame.get('frame_color'), '#000000'),
        'text_color': _color(frame.get('text_color', frame.get('frame_text_color', legacy_text_color)), '#FFFFFF'),
    }


def normalize_style_config(value=None, *, body_style='square', eye_style='square',
                            ball_style='square', fg_color='#000000',
                            bg_color='#ffffff', eye_color_outer=None,
                            eye_color_inner=None):
    """Return a safe, complete style config while accepting legacy fields."""
    source = value if isinstance(value, dict) else {}
    nested = source.get('style_config') if isinstance(source.get('style_config'), dict) else source
    patterns = nested.get('patterns') if isinstance(nested.get('patterns'), dict) else {}
    colors = nested.get('colors') if isinstance(nested.get('colors'), dict) else {}

    body = patterns.get('body', source.get('body_pattern', body_style))
    outer = patterns.get('outer_eye', source.get('outer_eye_style', eye_style))
    inner = patterns.get('inner_eye', source.get('inner_eye_style', ball_style))
    body = body if body in BODY_PATTERNS else 'square'
    outer = outer if outer in OUTER_EYE_STYLES else 'square'
    inner = inner if inner in INNER_EYE_STYLES else 'square'

    body_color = _color(colors.get('body', source.get('body_color', fg_color)), '#000000')
    outer_color = _color(colors.get('outer_eye', source.get('outer_eye_color', eye_color_outer or fg_color)), '#000000')
    inner_color = _color(colors.get('inner_eye', source.get('inner_eye_color', eye_color_inner or fg_color)), '#000000')
    background = _color(colors.get('background', source.get('background_color', bg_color)), '#FFFFFF')
    frame_source = nested if isinstance(nested.get('frame'), dict) else source

    return {
        'version': 1,
        'patterns': {'body': body, 'outer_eye': outer, 'inner_eye': inner},
        'colors': {
            'body': body_color,
            'outer_eye': outer_color,
            'inner_eye': inner_color,
            'background': background,
        },
        'frame': normalize_frame_config(
            frame_source,
            legacy_style=source.get('frame_style', 'none'),
            legacy_text=source.get('frame_text', ''),
            legacy_font=source.get('frame_font', 'Arial'),
            legacy_text_color=source.get('frame_text_color', '#FFFFFF'),
        ),
    }


def style_config_with_legacy_fields(config):
    """Expose normalized styles in both new and existing storage formats."""
    config = normalize_style_config(config)
    return {
        **config,
        'body_pattern': config['patterns']['body'],
        'outer_eye_style': config['patterns']['outer_eye'],
        'inner_eye_style': config['patterns']['inner_eye'],
        'body_color': config['colors']['body'],
        'outer_eye_color': config['colors']['outer_eye'],
        'inner_eye_color': config['colors']['inner_eye'],
        'background_color': config['colors']['background'],
    }
