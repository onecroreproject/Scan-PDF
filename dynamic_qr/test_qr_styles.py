import os
import re

from django.test import SimpleTestCase
from PIL import Image

from converter.utils import generate_qr_code
from .qr_styles import normalize_frame_config, normalize_style_config


class QRStyleRegistryTests(SimpleTestCase):
    def test_invalid_style_and_colors_use_safe_defaults(self):
        config = normalize_style_config({
            'patterns': {'body': 'not-real', 'outer_eye': 'circle', 'inner_eye': 'diamond'},
            'colors': {'body': '<svg>', 'outer_eye': '#123456', 'inner_eye': '#654321'},
        })

        self.assertEqual(config['patterns']['body'], 'square')
        self.assertEqual(config['patterns']['outer_eye'], 'circle')
        self.assertEqual(config['colors']['body'], '#000000')
        self.assertEqual(config['colors']['background'], '#FFFFFF')

    def test_expanded_styles_render_to_png_and_svg(self):
        config = {
            'style_config': {
                'patterns': {'body': 'sparkle', 'outer_eye': 'octagon', 'inner_eye': 'clover'},
                'colors': {
                    'body': '#112233',
                    'outer_eye': '#334455',
                    'inner_eye': '#556677',
                    'background': '#FFFFFF',
                },
            }
        }
        paths = []
        try:
            for output_format in ('png', 'svg'):
                path = generate_qr_code(
                    'https://scanpdf.com/r/style-test',
                    output_format=output_format,
                    design_options=config,
                )
                paths.append(path)
                self.assertTrue(os.path.exists(path))
                self.assertGreater(os.path.getsize(path), 100)
        finally:
            for path in paths:
                if os.path.exists(path):
                    os.remove(path)

    def test_frame_config_is_allow_listed_and_clamped(self):
        frame = normalize_frame_config({
            'type': 'bottom_bar',
            'text': 'x' * 80,
            'text_size': 99,
            'font': 'not-a-font',
            'frame_color': '#123456',
        })

        self.assertEqual(frame['type'], 'bottom_bar')
        self.assertEqual(len(frame['text']), 40)
        self.assertEqual(frame['text_size'], 40)
        self.assertEqual(frame['font'], 'Arial')
        self.assertEqual(frame['frame_color'], '#123456')

    def test_frame_is_present_in_png_and_svg_exports(self):
        config = {
            'style_config': {
                'patterns': {'body': 'square', 'outer_eye': 'square', 'inner_eye': 'square'},
                'colors': {'body': '#000000', 'outer_eye': '#000000', 'inner_eye': '#000000', 'background': '#FFFFFF'},
                'frame': {
                    'type': 'bottom_bar', 'text': 'SCAN ME', 'text_size': 18,
                    'font': 'Arial', 'use_custom_colors': True,
                    'frame_color': '#112233', 'text_color': '#FFFFFF',
                },
            }
        }
        paths = []
        try:
            for output_format in ('png', 'svg'):
                path = generate_qr_code('https://scanpdf.com/frame-test', output_format=output_format, design_options=config)
                paths.append(path)
                self.assertTrue(os.path.exists(path))
                if output_format == 'svg':
                    contents = open(path, encoding='utf-8').read()
                    self.assertIn('SCAN ME', contents)
                    self.assertIn('#112233', contents)
        finally:
            for path in paths:
                if os.path.exists(path):
                    os.remove(path)

    def test_text_frames_keep_exact_id_geometry_and_text_size(self):
        frame_ids = ('bottom_bar', 'speech_top', 'speech_bottom', 'ribbon_top', 'ribbon_bottom', 'scan_me_top', 'scan_me_bottom')
        paths = []
        try:
            for frame_id in frame_ids:
                path = generate_qr_code(
                    'https://scanpdf.com/frame-case',
                    output_format='svg',
                    design_options={'frame': {
                        'type': frame_id, 'text': 'SCAN ME', 'text_size': 33,
                        'font': 'Times New Roman', 'use_custom_colors': False,
                    }},
                )
                paths.append(path)
                svg = open(path, encoding='utf-8').read()
                self.assertIn('SCAN ME', svg)
                self.assertIn('font-size="33"', svg)
                self.assertIn('dominant-baseline="middle"', svg)
        finally:
            for path in paths:
                if os.path.exists(path):
                    os.remove(path)

    def test_circle_frame_expands_canvas_around_qr(self):
        path = generate_qr_code(
            'https://scanpdf.com/circle-case',
            output_format='svg',
            design_options={'frame': {'type': 'simple_circle'}},
        )
        try:
            svg = open(path, encoding='utf-8').read()
            viewbox = re.search(r'viewBox="0 0 (\d+) (\d+)"', svg)
            circle = re.search(r'<circle cx="([0-9.]+)" cy="([0-9.]+)" r="([0-9.]+)"', svg)
            self.assertIsNotNone(viewbox)
            self.assertIsNotNone(circle)
            self.assertEqual(viewbox.group(1), viewbox.group(2))
            self.assertGreater(float(circle.group(3)), 150)
        finally:
            if os.path.exists(path):
                os.remove(path)

    def test_png_frame_text_uses_selected_font_size(self):
        paths = []
        bounds = []
        try:
            for text_size in (20, 36):
                path = generate_qr_code(
                    'https://scanpdf.com/png-text-size',
                    output_format='png',
                    design_options={'style_config': {'frame': {
                        'type': 'top_bar', 'text': 'BALA', 'text_size': text_size,
                        'font': 'Arial', 'use_custom_colors': False,
                    }}},
                )
                paths.append(path)
                image = Image.open(path).convert('RGB')
                points = [
                    (x, y)
                    for y in range(8, 72)
                    for x in range(image.width // 2 - 120, image.width // 2 + 120)
                    if min(image.getpixel((x, y))) > 220
                ]
                bounds.append((max(x for x, _ in points) - min(x for x, _ in points), max(y for _, y in points) - min(y for _, y in points)))
            self.assertGreater(bounds[1][0], bounds[0][0])
            self.assertGreater(bounds[1][1], bounds[0][1])
        finally:
            for path in paths:
                if os.path.exists(path):
                    os.remove(path)
