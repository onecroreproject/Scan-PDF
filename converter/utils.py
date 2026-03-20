"""
Conversion utility functions for all supported file types.
Each function takes an input file path and returns the output file path.
"""
import os
import io
import tempfile
from pathlib import Path

from django.conf import settings


def ensure_media_dirs():
    """Ensure the media upload and output directories exist."""
    upload_dir = os.path.join(settings.MEDIA_ROOT, 'uploads')
    output_dir = os.path.join(settings.MEDIA_ROOT, 'outputs')
    os.makedirs(upload_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)
    return upload_dir, output_dir


def save_uploaded_file(uploaded_file):
    """Save an uploaded file to the media/uploads directory and return its path."""
    upload_dir, _ = ensure_media_dirs()
    file_path = os.path.join(upload_dir, uploaded_file.name)
    with open(file_path, 'wb+') as dest:
        for chunk in uploaded_file.chunks():
            dest.write(chunk)
    return file_path


def get_output_path(original_name, new_extension):
    """Generate an output file path with the new extension."""
    _, output_dir = ensure_media_dirs()
    base_name = Path(original_name).stem
    output_name = f"{base_name}.{new_extension}"
    return os.path.join(output_dir, output_name)


# ═══════════════════════════════════════════════════════════════
# 1. WORD (.docx) → PDF
# ═══════════════════════════════════════════════════════════════
def convert_word_to_pdf(input_path, original_name):
    """Convert a Word document (.docx) to PDF using python-docx + reportlab-style approach."""
    from docx import Document
    import fitz  # PyMuPDF

    output_path = get_output_path(original_name, 'pdf')

    doc = Document(input_path)
    
    # Create a PDF using PyMuPDF
    pdf_doc = fitz.open()
    
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        
        # Determine font size based on paragraph style
        font_size = 11
        is_bold = False
        style_name = para.style.name.lower() if para.style else ''
        
        if 'heading 1' in style_name:
            font_size = 24
            is_bold = True
        elif 'heading 2' in style_name:
            font_size = 20
            is_bold = True
        elif 'heading 3' in style_name:
            font_size = 16
            is_bold = True
        elif 'title' in style_name:
            font_size = 28
            is_bold = True
    
    # If no pages created yet, we need a different approach
    # Use a simple text-to-PDF method via PyMuPDF
    pdf_doc.close()
    
    # Use PyMuPDF Story for proper text rendering
    pdf_doc = fitz.open()
    
    # Extract all text from the document
    full_text_parts = []
    for para in doc.paragraphs:
        text = para.text
        if text.strip():
            style_name = para.style.name.lower() if para.style else ''
            if 'heading 1' in style_name or 'title' in style_name:
                full_text_parts.append(f'<h1>{text}</h1>')
            elif 'heading 2' in style_name:
                full_text_parts.append(f'<h2>{text}</h2>')
            elif 'heading 3' in style_name:
                full_text_parts.append(f'<h3>{text}</h3>')
            else:
                full_text_parts.append(f'<p>{text}</p>')
        else:
            full_text_parts.append('<br/>')
    
    # Also extract tables
    for table in doc.tables:
        table_html = '<table border="1" style="border-collapse: collapse; width: 100%;">'
        for row in table.rows:
            table_html += '<tr>'
            for cell in row.cells:
                table_html += f'<td style="padding: 4px; border: 1px solid #333;">{cell.text}</td>'
            table_html += '</tr>'
        table_html += '</table><br/>'
        full_text_parts.append(table_html)
    
    html_content = f"""
    <html>
    <head>
        <style>
            body {{ font-family: Helvetica, Arial, sans-serif; font-size: 11pt; line-height: 1.6; margin: 40px; }}
            h1 {{ font-size: 22pt; color: #1a1a2e; margin-bottom: 10px; }}
            h2 {{ font-size: 18pt; color: #16213e; margin-bottom: 8px; }}
            h3 {{ font-size: 14pt; color: #0f3460; margin-bottom: 6px; }}
            p {{ margin-bottom: 6px; color: #333; }}
            table {{ margin: 10px 0; }}
        </style>
    </head>
    <body>{''.join(full_text_parts)}</body>
    </html>
    """
    
    # Use WeasyPrint for HTML-to-PDF (most reliable on Windows)
    try:
        import weasyprint
        weasyprint.HTML(string=html_content).write_pdf(output_path)
    except Exception:
        # Fallback: use PyMuPDF simple text insertion
        pdf_doc = fitz.open()
        page = pdf_doc.new_page()
        y_position = 72
        
        for para in doc.paragraphs:
            text = para.text.strip()
            if not text:
                y_position += 12
                continue
            
            fontsize = 11
            style_name = para.style.name.lower() if para.style else ''
            if 'heading' in style_name or 'title' in style_name:
                fontsize = 16
            
            # Word wrap manually
            words = text.split()
            line = ""
            for word in words:
                test_line = f"{line} {word}".strip()
                if len(test_line) * fontsize * 0.5 > 470:  # approximate width
                    if y_position > 750:
                        page = pdf_doc.new_page()
                        y_position = 72
                    page.insert_text((72, y_position), line, fontsize=fontsize)
                    y_position += fontsize + 4
                    line = word
                else:
                    line = test_line
            
            if line:
                if y_position > 750:
                    page = pdf_doc.new_page()
                    y_position = 72
                page.insert_text((72, y_position), line, fontsize=fontsize)
                y_position += fontsize + 8
        
        pdf_doc.save(output_path)
        pdf_doc.close()
    
    return output_path


# ═══════════════════════════════════════════════════════════════
# 2. POWERPOINT (.pptx) → PDF
# ═══════════════════════════════════════════════════════════════
def _emu_to_px(emu):
    """Convert EMU to CSS pixels (96 DPI)."""
    if emu is None:
        return 0
    return emu / 914400 * 96


def _rgb_from_pptx_color(color_obj):
    """Try to extract an RGB hex string from a python-pptx color object."""
    try:
        if color_obj and color_obj.type is not None:
            rgb = color_obj.rgb  # RGBColor object
            return f'#{rgb}'
    except Exception:
        pass
    return None


def _build_run_html(run):
    """Convert a single python-pptx Run into an HTML span with inline styles."""
    import html as html_mod
    text = html_mod.escape(run.text)
    if not text:
        return ''

    styles = []
    font = run.font

    # Font size
    if font.size:
        styles.append(f'font-size:{font.size.pt}pt')

    # Bold / Italic / Underline
    if font.bold:
        styles.append('font-weight:bold')
    if font.italic:
        styles.append('font-style:italic')
    if font.underline:
        styles.append('text-decoration:underline')

    # Font color
    color_hex = _rgb_from_pptx_color(font.color)
    if color_hex:
        styles.append(f'color:{color_hex}')

    # Font family
    if font.name:
        styles.append(f"font-family:'{font.name}',Arial,sans-serif")

    style_attr = ';'.join(styles)
    return f'<span style="{style_attr}">{text}</span>'


def _build_paragraph_html(paragraph):
    """Convert a python-pptx Paragraph into an HTML <p> element."""
    runs_html = ''.join(_build_run_html(r) for r in paragraph.runs)
    if not runs_html.strip():
        return '<p style="margin:0;min-height:0.5em;">&nbsp;</p>'

    p_styles = ['margin:0 0 2px 0']

    # Alignment
    from pptx.enum.text import PP_ALIGN
    align_map = {
        PP_ALIGN.CENTER: 'center',
        PP_ALIGN.RIGHT: 'right',
        PP_ALIGN.JUSTIFY: 'justify',
    }
    if paragraph.alignment and paragraph.alignment in align_map:
        p_styles.append(f'text-align:{align_map[paragraph.alignment]}')

    # Line spacing
    if paragraph.line_spacing and hasattr(paragraph.line_spacing, 'pt'):
        p_styles.append(f'line-height:{paragraph.line_spacing.pt}pt')

    style_attr = ';'.join(p_styles)
    return f'<p style="{style_attr}">{runs_html}</p>'


def _extract_shape_fill_css(shape):
    """Try to get a CSS background from the shape's fill."""
    try:
        fill = shape.fill
        if fill and fill.type is not None:
            from pptx.enum.dml import MSO_THEME_COLOR
            # Solid fill
            if fill.type == 1:  # MSO_FILL_TYPE.SOLID
                rgb = fill.fore_color.rgb
                return f'background-color:#{rgb};'
    except Exception:
        pass
    return ''


def convert_pptx_to_pdf(input_path, original_name):
    """Convert a PowerPoint presentation (.pptx) to PDF with high-quality rendering.

    Strategy: build a multi-page HTML document where every slide is a fixed-size
    CSS page with shapes positioned absolutely at their original coordinates,
    then render to PDF with WeasyPrint for best quality.
    """
    from pptx import Presentation
    from pptx.util import Emu
    import base64
    import html as html_mod

    output_path = get_output_path(original_name, 'pdf')

    prs = Presentation(input_path)
    slide_w_emu = prs.slide_width or Emu(9144000)   # 10 in
    slide_h_emu = prs.slide_height or Emu(6858000)  # 7.5 in
    slide_w_px = _emu_to_px(slide_w_emu)
    slide_h_px = _emu_to_px(slide_h_emu)

    # Build CSS for each page to match the slide dimensions exactly
    page_css = f"""
    @page {{
        size: {slide_w_px}px {slide_h_px}px;
        margin: 0;
    }}
    * {{ box-sizing: border-box; }}
    body {{ margin:0; padding:0; font-family: Arial, Helvetica, sans-serif; }}
    .slide {{
        width: {slide_w_px}px;
        height: {slide_h_px}px;
        position: relative;
        overflow: hidden;
        background: #ffffff;
        page-break-after: always;
    }}
    .slide:last-child {{ page-break-after: auto; }}
    .shape {{
        position: absolute;
        overflow: hidden;
        word-wrap: break-word;
    }}
    .shape-text {{
        padding: 4px 8px;
    }}
    table.pptx-table {{
        border-collapse: collapse;
        width: 100%;
        height: 100%;
    }}
    table.pptx-table td {{
        border: 1px solid #bbb;
        padding: 4px 6px;
        font-size: 10pt;
        vertical-align: middle;
    }}
    """

    slides_html = []

    for slide_idx, slide in enumerate(prs.slides):
        shapes_html = []

        # --- try to get slide background colour ---
        slide_bg = '#ffffff'
        try:
            bg = slide.background
            if bg.fill and bg.fill.type is not None and bg.fill.type == 1:
                slide_bg = f'#{bg.fill.fore_color.rgb}'
        except Exception:
            pass

        # Sort shapes by their z-order (shape_id) so layering is correct
        sorted_shapes = sorted(slide.shapes, key=lambda s: s.shape_id)

        for shape in sorted_shapes:
            left = _emu_to_px(shape.left)
            top = _emu_to_px(shape.top)
            width = _emu_to_px(shape.width)
            height = _emu_to_px(shape.height)

            shape_style = (
                f'left:{left}px;top:{top}px;'
                f'width:{width}px;height:{height}px;'
            )

            # Shape fill
            fill_css = _extract_shape_fill_css(shape)
            if fill_css:
                shape_style += fill_css

            # Rotation
            if shape.rotation:
                shape_style += f'transform:rotate({shape.rotation}deg);'

            inner_html = ''

            # ── Image shapes ────────────────────────────
            if shape.shape_type and shape.shape_type == 13:  # MSO_SHAPE_TYPE.PICTURE
                try:
                    image = shape.image
                    blob = image.blob
                    content_type = image.content_type or 'image/png'
                    b64 = base64.b64encode(blob).decode('ascii')
                    inner_html = (
                        f'<img src="data:{content_type};base64,{b64}" '
                        f'style="width:100%;height:100%;object-fit:contain;" />'
                    )
                except Exception:
                    pass

            # ── Tables ──────────────────────────────────
            elif shape.has_table:
                table = shape.table
                rows_html = []
                for row in table.rows:
                    cells_html = []
                    for cell in row.cells:
                        cell_text = html_mod.escape(cell.text)
                        cell_bg = ''
                        try:
                            if cell.fill and cell.fill.type is not None and cell.fill.type == 1:
                                cell_bg = f'background-color:#{cell.fill.fore_color.rgb};'
                        except Exception:
                            pass
                        cells_html.append(
                            f'<td style="{cell_bg}">{cell_text}</td>'
                        )
                    rows_html.append('<tr>' + ''.join(cells_html) + '</tr>')
                inner_html = (
                    '<table class="pptx-table">'
                    + ''.join(rows_html)
                    + '</table>'
                )

            # ── Text frames ─────────────────────────────
            elif shape.has_text_frame:
                tf = shape.text_frame
                paras_html = ''.join(
                    _build_paragraph_html(p) for p in tf.paragraphs
                )

                # Vertical alignment
                vert_align_css = 'justify-content:flex-start;'
                try:
                    from pptx.enum.text import MSO_ANCHOR
                    if tf.word_wrap is not None:
                        pass  # just accessing to ensure tf is valid
                    anchor = tf.paragraphs[0]  # dummy access
                    # Use the text frame's anchor property if available
                    if hasattr(tf, '_txBody'):
                        anchor_val = tf._txBody.bodyPr.get('anchor', 't')
                        if anchor_val == 'ctr':
                            vert_align_css = 'justify-content:center;'
                        elif anchor_val == 'b':
                            vert_align_css = 'justify-content:flex-end;'
                except Exception:
                    pass

                inner_html = (
                    f'<div class="shape-text" style="display:flex;flex-direction:column;'
                    f'height:100%;{vert_align_css}">'
                    f'{paras_html}</div>'
                )

            # Only add shape if it has content
            if inner_html:
                shapes_html.append(
                    f'<div class="shape" style="{shape_style}">{inner_html}</div>'
                )

        slide_div = (
            f'<div class="slide" style="background:{slide_bg};">'
            + ''.join(shapes_html)
            + '</div>'
        )
        slides_html.append(slide_div)

    if not slides_html:
        slides_html.append(
            f'<div class="slide"><p style="padding:40px;font-size:14pt;">'
            f'Empty presentation – no content to convert.</p></div>'
        )

    full_html = (
        '<!DOCTYPE html><html><head><meta charset="utf-8">'
        f'<style>{page_css}</style></head><body>'
        + ''.join(slides_html)
        + '</body></html>'
    )

    # ── Primary path: WeasyPrint (best quality) ─────────
    try:
        import weasyprint
        weasyprint.HTML(string=full_html).write_pdf(output_path)
        return output_path
    except Exception:
        pass

    # ── Fallback: render HTML pages as images with PyMuPDF ─
    try:
        import fitz
        # Write HTML to a temp file so PyMuPDF can attempt to open it
        import tempfile
        tmp_html = tempfile.NamedTemporaryFile(
            suffix='.html', delete=False, mode='w', encoding='utf-8'
        )
        tmp_html.write(full_html)
        tmp_html.close()

        # Fallback: manual text extraction when neither renderer works
        pdf_doc = fitz.open()

        page_w_pt = slide_w_emu / 914400 * 72
        page_h_pt = slide_h_emu / 914400 * 72

        for slide_idx, slide in enumerate(prs.slides):
            page = pdf_doc.new_page(width=page_w_pt, height=page_h_pt)
            page.draw_rect(
                fitz.Rect(0, 0, page_w_pt, page_h_pt),
                color=(1, 1, 1), fill=(1, 1, 1),
            )

            y_cursor = 36  # running y position for sequential text

            for shape in sorted(slide.shapes, key=lambda s: s.shape_id):
                s_top_pt = (shape.top / 914400 * 72) if shape.top else y_cursor
                s_left_pt = (shape.left / 914400 * 72) if shape.left else 36
                s_width_pt = (shape.width / 914400 * 72) if shape.width else 500

                if shape.has_text_frame:
                    cur_y = s_top_pt + 14  # small inner padding
                    for para in shape.text_frame.paragraphs:
                        text = para.text.strip()
                        if not text:
                            cur_y += 10
                            continue

                        fontsize = 12
                        is_bold = False
                        for run in para.runs:
                            if run.font.size:
                                fontsize = run.font.size.pt
                                break
                            if run.font.bold:
                                is_bold = True
                        fontsize = max(7, min(fontsize, 48))

                        # Simple word-wrap
                        max_chars = max(int(s_width_pt / (fontsize * 0.52)), 10)
                        words = text.split()
                        line = ''
                        for word in words:
                            test = f'{line} {word}'.strip()
                            if len(test) > max_chars and line:
                                if cur_y > page_h_pt - 20:
                                    page = pdf_doc.new_page(width=page_w_pt, height=page_h_pt)
                                    page.draw_rect(fitz.Rect(0, 0, page_w_pt, page_h_pt), fill=(1, 1, 1))
                                    cur_y = 36
                                page.insert_text(
                                    (s_left_pt + 4, cur_y), line,
                                    fontsize=fontsize, color=(0.1, 0.1, 0.1),
                                )
                                cur_y += fontsize + 3
                                line = word
                            else:
                                line = test
                        if line:
                            if cur_y > page_h_pt - 20:
                                page = pdf_doc.new_page(width=page_w_pt, height=page_h_pt)
                                page.draw_rect(fitz.Rect(0, 0, page_w_pt, page_h_pt), fill=(1, 1, 1))
                                cur_y = 36
                            page.insert_text(
                                (s_left_pt + 4, cur_y), line,
                                fontsize=fontsize, color=(0.1, 0.1, 0.1),
                            )
                            cur_y += fontsize + 5

                elif shape.has_table:
                    table = shape.table
                    cur_y = s_top_pt + 14
                    for row in table.rows:
                        col_x = s_left_pt + 4
                        col_w = max(int(s_width_pt / max(len(row.cells), 1)), 40)
                        for cell in row.cells:
                            ct = cell.text.strip()[:30]
                            if ct:
                                page.insert_text(
                                    (col_x, cur_y), ct,
                                    fontsize=9, color=(0.15, 0.15, 0.15),
                                )
                            col_x += col_w
                        cur_y += 16

        if len(pdf_doc) == 0:
            p = pdf_doc.new_page()
            p.insert_text((72, 72), 'Empty presentation.', fontsize=14)

        pdf_doc.save(output_path)
        pdf_doc.close()

        # Clean temp file
        try:
            os.remove(tmp_html.name)
        except OSError:
            pass

        return output_path
    except Exception as e:
        raise Exception(f'Failed to convert PPTX to PDF: {str(e)}')


# ═══════════════════════════════════════════════════════════════
# 3. EXCEL (.xlsx) → PDF
# ═══════════════════════════════════════════════════════════════
def _openpyxl_color_to_hex(color):
    """Extract hex colour from an openpyxl Color object. Returns None on failure."""
    if color is None:
        return None
    try:
        if color.type == 'rgb' and color.rgb and color.rgb != '00000000':
            rgb = color.rgb
            # openpyxl stores ARGB – skip first two chars (alpha)
            if len(rgb) == 8:
                rgb = rgb[2:]
            return f'#{rgb}'
        if color.type == 'indexed':
            # A few common indexed colours
            indexed_map = {
                0: '#000000', 1: '#FFFFFF', 2: '#FF0000', 3: '#00FF00',
                4: '#0000FF', 5: '#FFFF00', 6: '#FF00FF', 7: '#00FFFF',
                8: '#000000', 9: '#FFFFFF', 10: '#FF0000', 11: '#00FF00',
                22: '#C0C0C0', 55: '#808080', 64: None,
            }
            return indexed_map.get(color.indexed, None)
        if color.type == 'theme':
            # Theme colours are tricky – skip and let the fallback handle it
            return None
    except Exception:
        pass
    return None


def convert_excel_to_pdf(input_path, original_name):
    """Convert an Excel spreadsheet (.xlsx) to PDF with high-quality formatting.

    Strategy: read cell-level formatting via openpyxl (background colours, fonts,
    alignment, borders, merged cells, column widths) and build a richly-styled
    HTML table that mirrors the spreadsheet's visual appearance, then render to
    PDF with WeasyPrint.
    """
    import html as html_mod
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter

    output_path = get_output_path(original_name, 'pdf')

    wb = load_workbook(input_path, data_only=True)

    # Decide page orientation: landscape when any sheet has many columns
    use_landscape = False
    for ws in wb.worksheets:
        if ws.max_column and ws.max_column > 6:
            use_landscape = True
            break

    page_size = '297mm 210mm' if use_landscape else '210mm 297mm'  # A4

    page_css = f"""
    @page {{
        size: {page_size};
        margin: 15mm 12mm;
    }}
    * {{ box-sizing: border-box; }}
    body {{
        margin: 0; padding: 0;
        font-family: 'Segoe UI', Arial, Helvetica, sans-serif;
        font-size: 9pt;
        color: #1e293b;
    }}
    .sheet-section {{
        page-break-after: always;
    }}
    .sheet-section:last-child {{
        page-break-after: auto;
    }}
    .sheet-title {{
        font-size: 15pt;
        font-weight: 700;
        color: #1e293b;
        margin: 0 0 10px 0;
        padding-bottom: 6px;
        border-bottom: 3px solid #4f46e5;
    }}
    .sheet-meta {{
        font-size: 8pt;
        color: #94a3b8;
        margin-bottom: 12px;
    }}
    table {{
        border-collapse: collapse;
        width: 100%;
        margin-bottom: 20px;
        table-layout: auto;
    }}
    th {{
        background-color: #4f46e5;
        color: #ffffff;
        font-weight: 600;
        font-size: 8.5pt;
        padding: 7px 10px;
        text-align: left;
        border: 1px solid #4338ca;
        white-space: nowrap;
    }}
    td {{
        padding: 5px 10px;
        border: 1px solid #e2e8f0;
        font-size: 8.5pt;
        vertical-align: middle;
        word-wrap: break-word;
    }}
    tr:nth-child(even) td {{
        background-color: #f8fafc;
    }}
    """

    sheets_html = []

    for ws in wb.worksheets:
        if ws.max_row is None or ws.max_row < 1:
            continue
        if ws.max_column is None or ws.max_column < 1:
            continue

        # Gather merged-cell spans
        merged_ranges = {}
        skip_cells = set()
        for merged in ws.merged_cells.ranges:
            min_row, min_col = merged.min_row, merged.min_col
            max_row, max_col = merged.max_row, merged.max_col
            rowspan = max_row - min_row + 1
            colspan = max_col - min_col + 1
            merged_ranges[(min_row, min_col)] = (rowspan, colspan)
            for r in range(min_row, max_row + 1):
                for c in range(min_col, max_col + 1):
                    if (r, c) != (min_row, min_col):
                        skip_cells.add((r, c))

        # Build column width hints
        col_widths = {}
        for col_idx in range(1, ws.max_column + 1):
            letter = get_column_letter(col_idx)
            dim = ws.column_dimensions.get(letter)
            if dim and dim.width:
                col_widths[col_idx] = max(dim.width * 7, 40)  # approximate px

        # Build rows
        rows_html = []
        for row_idx in range(1, ws.max_row + 1):
            cells_html = []
            for col_idx in range(1, ws.max_column + 1):
                if (row_idx, col_idx) in skip_cells:
                    continue

                cell = ws.cell(row=row_idx, column=col_idx)
                value = cell.value
                if value is None:
                    display = ''
                else:
                    # Apply number format for common patterns
                    try:
                        nf = cell.number_format
                        if nf and nf != 'General' and isinstance(value, (int, float)):
                            if '%' in nf:
                                display = f'{value * 100:.1f}%'
                            elif '0.00' in nf:
                                display = f'{value:,.2f}'
                            elif '0.0' in nf:
                                display = f'{value:,.1f}'
                            elif '#,##0' in nf:
                                display = f'{value:,.0f}'
                            else:
                                display = html_mod.escape(str(value))
                        else:
                            display = html_mod.escape(str(value))
                    except Exception:
                        display = html_mod.escape(str(value))

                # Build inline styles from cell formatting
                styles = []

                # Background colour
                if cell.fill and cell.fill.fgColor:
                    bg = _openpyxl_color_to_hex(cell.fill.fgColor)
                    if bg and bg.lower() != '#ffffff' and bg.lower() != '#000000':
                        styles.append(f'background-color:{bg}')

                # Font properties
                font = cell.font
                if font:
                    if font.bold:
                        styles.append('font-weight:bold')
                    if font.italic:
                        styles.append('font-style:italic')
                    if font.underline and font.underline != 'none':
                        styles.append('text-decoration:underline')
                    if font.size:
                        styles.append(f'font-size:{font.size}pt')
                    fc = _openpyxl_color_to_hex(font.color)
                    if fc:
                        styles.append(f'color:{fc}')
                    if font.name and font.name != 'Calibri':
                        styles.append(f"font-family:'{font.name}',Arial,sans-serif")

                # Alignment
                alignment = cell.alignment
                if alignment:
                    ha = alignment.horizontal
                    if ha:
                        align_map = {'center': 'center', 'right': 'right',
                                     'left': 'left', 'justify': 'justify'}
                        if ha in align_map:
                            styles.append(f'text-align:{align_map[ha]}')
                    va = alignment.vertical
                    if va:
                        va_map = {'center': 'middle', 'top': 'top', 'bottom': 'bottom'}
                        if va in va_map:
                            styles.append(f'vertical-align:{va_map[va]}')
                    if alignment.wrap_text:
                        styles.append('white-space:normal;word-wrap:break-word')

                # Column width
                if col_idx in col_widths:
                    styles.append(f'min-width:{col_widths[col_idx]}px')

                style_attr = ';'.join(styles)

                # Use <th> for the first row (header), <td> for data
                tag = 'th' if row_idx == 1 else 'td'

                # Merged cell attributes
                span_attrs = ''
                if (row_idx, col_idx) in merged_ranges:
                    rspan, cspan = merged_ranges[(row_idx, col_idx)]
                    if rspan > 1:
                        span_attrs += f' rowspan="{rspan}"'
                    if cspan > 1:
                        span_attrs += f' colspan="{cspan}"'

                cells_html.append(
                    f'<{tag}{span_attrs} style="{style_attr}">{display}</{tag}>'
                )

            if cells_html:
                rows_html.append('<tr>' + ''.join(cells_html) + '</tr>')

        if not rows_html:
            continue

        sheet_html = f"""
        <div class="sheet-section">
            <h2 class="sheet-title">{html_mod.escape(ws.title)}</h2>
            <p class="sheet-meta">{ws.max_row - 1} rows × {ws.max_column} columns</p>
            <table>
                {''.join(rows_html)}
            </table>
        </div>
        """
        sheets_html.append(sheet_html)

    if not sheets_html:
        sheets_html.append(
            '<div class="sheet-section"><p style="font-size:13pt;padding:40px;">'
            'The spreadsheet contains no data to convert.</p></div>'
        )

    full_html = (
        '<!DOCTYPE html><html><head><meta charset="utf-8">'
        f'<style>{page_css}</style></head><body>'
        + ''.join(sheets_html)
        + '</body></html>'
    )

    # ── Primary path: WeasyPrint (best quality) ─────────
    try:
        import weasyprint
        weasyprint.HTML(string=full_html).write_pdf(output_path)
        return output_path
    except Exception:
        pass

    # ── Fallback: PyMuPDF text-based rendering ──────────
    try:
        import fitz
        import pandas as pd

        pdf_doc = fitz.open()

        for ws in wb.worksheets:
            if ws.max_row is None or ws.max_row < 1:
                continue

            page = pdf_doc.new_page(width=842 if use_landscape else 595,
                                    height=595 if use_landscape else 842)
            margin = 40
            usable_w = page.rect.width - 2 * margin
            usable_h = page.rect.height - 2 * margin
            y = margin

            # Sheet title
            page.insert_text((margin, y + 14), ws.title, fontsize=14,
                           color=(0.31, 0.27, 0.89))
            y += 28

            # Draw a line under the title
            page.draw_line((margin, y), (margin + usable_w, y),
                          color=(0.31, 0.27, 0.89), width=1.5)
            y += 12

            max_col = min(ws.max_column or 1, 20)  # Cap columns
            col_w = usable_w / max_col
            row_h = 16

            for row_idx in range(1, (ws.max_row or 0) + 1):
                if y + row_h > margin + usable_h:
                    page = pdf_doc.new_page(
                        width=842 if use_landscape else 595,
                        height=595 if use_landscape else 842
                    )
                    y = margin

                for col_idx in range(1, max_col + 1):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    val = str(cell.value) if cell.value is not None else ''
                    x = margin + (col_idx - 1) * col_w

                    # Truncate to fit column
                    max_chars = max(int(col_w / 5.5), 4)
                    display = val[:max_chars]

                    fontsize = 8
                    color = (0.12, 0.12, 0.12)

                    if row_idx == 1:
                        fontsize = 8.5
                        color = (0.31, 0.27, 0.89)
                        # Draw header cell background
                        page.draw_rect(
                            fitz.Rect(x, y - 2, x + col_w, y + row_h - 2),
                            fill=(0.93, 0.93, 0.97),
                            color=(0.85, 0.85, 0.9),
                            width=0.3,
                        )

                    try:
                        page.insert_text((x + 3, y + 10), display,
                                        fontsize=fontsize, color=color)
                    except Exception:
                        pass

                    # Draw cell border
                    page.draw_rect(
                        fitz.Rect(x, y - 2, x + col_w, y + row_h - 2),
                        color=(0.88, 0.88, 0.88), width=0.3,
                    )

                y += row_h

        if len(pdf_doc) == 0:
            p = pdf_doc.new_page()
            p.insert_text((72, 72), 'Empty spreadsheet.', fontsize=14)

        pdf_doc.save(output_path)
        pdf_doc.close()
        return output_path

    except Exception as e:
        raise Exception(f'Failed to convert Excel to PDF: {str(e)}')


# ═══════════════════════════════════════════════════════════════
# 4. HTML → PDF
# ═══════════════════════════════════════════════════════════════
def convert_html_to_pdf(input_path, original_name):
    """Convert an HTML file to PDF."""
    output_path = get_output_path(original_name, 'pdf')
    
    with open(input_path, 'r', encoding='utf-8', errors='ignore') as f:
        html_content = f.read()
    
    # Try WeasyPrint first
    try:
        import weasyprint
        weasyprint.HTML(string=html_content).write_pdf(output_path)
        return output_path
    except Exception:
        pass
    
    # Fallback: PyMuPDF
    try:
        import fitz
        import re
        
        # Strip HTML tags for plain text extraction
        clean_text = re.sub(r'<[^>]+>', '\n', html_content)
        clean_text = re.sub(r'\n{3,}', '\n\n', clean_text)
        lines = [line.strip() for line in clean_text.split('\n') if line.strip()]
        
        pdf_doc = fitz.open()
        page = pdf_doc.new_page()
        y = 72
        
        for line in lines:
            if y > 750:
                page = pdf_doc.new_page()
                y = 72
            
            # Word wrap
            words = line.split()
            current_line = ""
            for word in words:
                test = f"{current_line} {word}".strip()
                if len(test) * 5.5 > 470:
                    page.insert_text((72, y), current_line, fontsize=11)
                    y += 16
                    current_line = word
                    if y > 750:
                        page = pdf_doc.new_page()
                        y = 72
                else:
                    current_line = test
            
            if current_line:
                page.insert_text((72, y), current_line, fontsize=11)
                y += 18
        
        pdf_doc.save(output_path)
        pdf_doc.close()
        return output_path
    except Exception as e:
        raise Exception(f"Failed to convert HTML to PDF: {str(e)}")


# ═══════════════════════════════════════════════════════════════
# 5. PDF → IMAGE (JPG/PNG)
# ═══════════════════════════════════════════════════════════════
def convert_pdf_to_image(input_path, original_name, image_format='png'):
    """Convert a PDF to images (one image per page). Returns a zip if multiple pages."""
    import fitz  # PyMuPDF
    import zipfile

    _, output_dir = ensure_media_dirs()
    base_name = Path(original_name).stem
    
    pdf_doc = fitz.open(input_path)
    num_pages = len(pdf_doc)
    
    if num_pages == 0:
        raise Exception("PDF has no pages to convert.")
    
    if num_pages == 1:
        # Single page: return a single image
        page = pdf_doc[0]
        # High quality rendering (2x zoom)
        mat = fitz.Matrix(2.0, 2.0)
        pix = page.get_pixmap(matrix=mat)
        
        output_path = os.path.join(output_dir, f"{base_name}.{image_format}")
        
        if image_format.lower() == 'jpg' or image_format.lower() == 'jpeg':
            # PyMuPDF saves as PNG by default, convert via Pillow
            from PIL import Image
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            img.save(output_path, "JPEG", quality=95)
        else:
            pix.save(output_path)
        
        pdf_doc.close()
        return output_path
    else:
        # Multiple pages: create a ZIP archive
        zip_path = os.path.join(output_dir, f"{base_name}_images.zip")
        
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for page_num in range(num_pages):
                page = pdf_doc[page_num]
                mat = fitz.Matrix(2.0, 2.0)
                pix = page.get_pixmap(matrix=mat)
                
                img_filename = f"{base_name}_page_{page_num + 1}.{image_format}"
                img_path = os.path.join(output_dir, img_filename)
                
                if image_format.lower() in ('jpg', 'jpeg'):
                    from PIL import Image
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    img.save(img_path, "JPEG", quality=95)
                else:
                    pix.save(img_path)
                
                zipf.write(img_path, img_filename)
                
                # Clean up individual image
                try:
                    os.remove(img_path)
                except OSError:
                    pass
        
        pdf_doc.close()
        return zip_path


# ═══════════════════════════════════════════════════════════════
# 6. PDF → WORD (.docx)
# ═══════════════════════════════════════════════════════════════
def convert_pdf_to_word(input_path, original_name):
    """Convert a PDF file to a Word document (.docx).

    Primary: pdf2docx (preserves layout, images, tables).
    Fallback: PyMuPDF text extraction into a styled python-docx document.
    """
    output_path = get_output_path(original_name, 'docx')

    # ── Primary: pdf2docx ───────────────────────────────
    try:
        from pdf2docx import Converter
        cv = Converter(input_path)
        cv.convert(output_path)
        cv.close()
        return output_path
    except Exception:
        pass

    # ── Fallback: PyMuPDF + python-docx ─────────────────
    try:
        import fitz
        from docx import Document
        from docx.shared import Pt, Inches, RGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH

        pdf = fitz.open(input_path)
        doc = Document()

        # Set default style
        style = doc.styles['Normal']
        style.font.name = 'Calibri'
        style.font.size = Pt(11)
        style.paragraph_format.space_after = Pt(4)

        for page_idx in range(len(pdf)):
            page = pdf[page_idx]

            if page_idx > 0:
                doc.add_page_break()

            # Add page header
            header_para = doc.add_paragraph()
            header_run = header_para.add_run(f'— Page {page_idx + 1} —')
            header_run.font.size = Pt(8)
            header_run.font.color.rgb = RGBColor(148, 163, 184)
            header_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            header_para.paragraph_format.space_after = Pt(12)

            # Extract text blocks with positions
            blocks = page.get_text("dict")["blocks"]

            for block in blocks:
                if block["type"] == 0:  # Text block
                    for line in block.get("lines", []):
                        para = doc.add_paragraph()
                        for span in line.get("spans", []):
                            text = span.get("text", "")
                            if not text.strip():
                                continue

                            run = para.add_run(text)

                            # Font size
                            size = span.get("size", 11)
                            run.font.size = Pt(max(6, min(size, 36)))

                            # Bold / Italic detection from flags
                            flags = span.get("flags", 0)
                            if flags & 2 ** 4:  # bold flag
                                run.bold = True
                            if flags & 2 ** 1:  # italic flag
                                run.italic = True

                            # Font colour
                            color_int = span.get("color", 0)
                            if color_int and color_int != 0:
                                r = (color_int >> 16) & 0xFF
                                g = (color_int >> 8) & 0xFF
                                b = color_int & 0xFF
                                run.font.color.rgb = RGBColor(r, g, b)

                            # Font family
                            font_name = span.get("font", "")
                            if font_name:
                                clean = font_name.split("+")[-1].split("-")[0]
                                run.font.name = clean

                elif block["type"] == 1:  # Image block
                    try:
                        img_data = block.get("image")
                        if img_data:
                            img_stream = io.BytesIO(img_data)
                            doc.add_picture(img_stream, width=Inches(5))
                    except Exception:
                        pass

        doc.save(output_path)
        pdf.close()
        return output_path

    except Exception as e:
        raise Exception(f'Failed to convert PDF to Word: {str(e)}')


# ═══════════════════════════════════════════════════════════════
# 7. PDF → POWERPOINT (.pptx)
# ═══════════════════════════════════════════════════════════════
def convert_pdf_to_pptx(input_path, original_name):
    """Convert a PDF file to a PowerPoint presentation (.pptx).

    Extracts each page as an image + text overlay into individual slides
    for maximum visual fidelity.
    """
    output_path = get_output_path(original_name, 'pptx')

    try:
        import fitz
        from pptx import Presentation
        from pptx.util import Inches, Pt, Emu
        from pptx.dml.color import RGBColor as PptxRGBColor
        from pptx.enum.text import PP_ALIGN

        pdf = fitz.open(input_path)
        prs = Presentation()

        # Set slide dimensions to match PDF aspect ratio
        first_page = pdf[0]
        pdf_w = first_page.rect.width
        pdf_h = first_page.rect.height

        # Scale to standard slide width (10 inches)
        slide_w_in = 10
        scale_factor = slide_w_in / (pdf_w / 72)
        slide_h_in = (pdf_h / 72) * scale_factor

        prs.slide_width = Inches(slide_w_in)
        prs.slide_height = Inches(slide_h_in)

        blank_layout = prs.slide_layouts[6]  # Blank layout

        for page_idx in range(len(pdf)):
            page = pdf[page_idx]
            slide = prs.slides.add_slide(blank_layout)

            # ── Render page as background image ─────────────
            mat = fitz.Matrix(2.5, 2.5)  # High resolution
            pix = page.get_pixmap(matrix=mat)
            img_data = pix.tobytes("png")
            img_stream = io.BytesIO(img_data)

            slide.shapes.add_picture(
                img_stream,
                Inches(0), Inches(0),
                Inches(slide_w_in), Inches(slide_h_in)
            )

            # ── Overlay extracted text for copy-paste support ─
            blocks = page.get_text("dict")["blocks"]
            scale_x = slide_w_in / (pdf_w / 72)
            scale_y = slide_h_in / (pdf_h / 72)

            for block in blocks:
                if block["type"] != 0:
                    continue

                bbox = block.get("bbox", [0, 0, 0, 0])
                bx0 = bbox[0] / 72 * scale_x
                by0 = bbox[1] / 72 * scale_y
                bw = (bbox[2] - bbox[0]) / 72 * scale_x
                bh = (bbox[3] - bbox[1]) / 72 * scale_y

                if bw < 0.1 or bh < 0.1:
                    continue

                # Create a transparent text box
                txBox = slide.shapes.add_textbox(
                    Inches(bx0), Inches(by0),
                    Inches(bw), Inches(bh)
                )
                tf = txBox.text_frame
                tf.word_wrap = True

                first_para = True
                for line in block.get("lines", []):
                    line_text = ""
                    font_size = 10
                    for span in line.get("spans", []):
                        line_text += span.get("text", "")
                        if span.get("size"):
                            font_size = span["size"]

                    if not line_text.strip():
                        continue

                    if first_para:
                        para = tf.paragraphs[0]
                        first_para = False
                    else:
                        para = tf.add_paragraph()

                    run = para.add_run()
                    run.text = line_text
                    run.font.size = Pt(max(5, min(font_size * scale_x * 0.7, 36)))
                    # Make text fully transparent so it doesn't visually interfere
                    # but is still selectable
                    run.font.color.rgb = PptxRGBColor(255, 255, 255)

                # Make text box fill transparent
                txBox.fill.background()

        prs.save(output_path)
        pdf.close()
        return output_path

    except Exception as e:
        raise Exception(f'Failed to convert PDF to PowerPoint: {str(e)}')


# ═══════════════════════════════════════════════════════════════
# 8. PDF → EXCEL (.xlsx)
# ═══════════════════════════════════════════════════════════════
def convert_pdf_to_excel(input_path, original_name):
    """Convert a PDF file to an Excel workbook (.xlsx).

    Primary: pdfplumber for accurate table detection and extraction.
    Fallback: line-based text extraction into columns when no tables found.
    """
    output_path = get_output_path(original_name, 'xlsx')

    try:
        import pdfplumber
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter

        wb = Workbook()
        wb.remove(wb.active)

        # Styling
        header_font = Font(name='Calibri', bold=True, size=10, color='FFFFFF')
        header_fill = PatternFill(start_color='4F46E5', end_color='4F46E5', fill_type='solid')
        header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell_font = Font(name='Calibri', size=10)
        cell_align = Alignment(vertical='center', wrap_text=True)
        thin_border = Border(
            left=Side(style='thin', color='E2E8F0'),
            right=Side(style='thin', color='E2E8F0'),
            top=Side(style='thin', color='E2E8F0'),
            bottom=Side(style='thin', color='E2E8F0'),
        )
        alt_fill = PatternFill(start_color='F8FAFC', end_color='F8FAFC', fill_type='solid')

        with pdfplumber.open(input_path) as pdf:
            table_counter = 0
            for page_idx, page in enumerate(pdf.pages):
                # Extract tables with pdfplumber's superior detection
                tables = page.extract_tables({
                    "vertical_strategy": "lines_strict",
                    "horizontal_strategy": "lines_strict",
                })

                # Fallback: try text-based strategy if strict found nothing
                if not tables:
                    tables = page.extract_tables({
                        "vertical_strategy": "text",
                        "horizontal_strategy": "text",
                        "snap_x_tolerance": 5,
                        "snap_y_tolerance": 5,
                        "join_x_tolerance": 5,
                        "join_y_tolerance": 5,
                    })

                if tables:
                    for tbl_idx, table_data in enumerate(tables):
                        if not table_data or len(table_data) == 0:
                            continue
                        table_counter += 1
                        sheet_name = f'Table {table_counter}'
                        if len(sheet_name) > 31:
                            sheet_name = sheet_name[:31]
                        ws = wb.create_sheet(title=sheet_name)

                        for row_idx, row in enumerate(table_data):
                            if row is None:
                                continue
                            for col_idx, cell_val in enumerate(row):
                                cell = ws.cell(
                                    row=row_idx + 1,
                                    column=col_idx + 1,
                                    value=(cell_val or '').strip()
                                )
                                cell.border = thin_border

                                # Try to convert numeric strings
                                if cell.value:
                                    cleaned = cell.value.replace(',', '').strip()
                                    try:
                                        if '.' in cleaned:
                                            cell.value = float(cleaned)
                                        else:
                                            cell.value = int(cleaned)
                                    except (ValueError, TypeError):
                                        pass

                                if row_idx == 0:
                                    cell.font = header_font
                                    cell.fill = header_fill
                                    cell.alignment = header_align
                                else:
                                    cell.font = cell_font
                                    cell.alignment = cell_align
                                    if row_idx % 2 == 0:
                                        cell.fill = alt_fill

                        # Auto-fit column widths
                        max_col = max(len(r) for r in table_data if r) if table_data else 0
                        for col_idx in range(1, max_col + 1):
                            max_len = 8
                            col_letter = get_column_letter(col_idx)
                            for row_idx in range(1, len(table_data) + 1):
                                val = ws.cell(row=row_idx, column=col_idx).value
                                if val is not None:
                                    max_len = max(max_len, min(len(str(val)) + 2, 60))
                            ws.column_dimensions[col_letter].width = max_len

                        ws.freeze_panes = 'A2'
                else:
                    # No tables found – extract text line by line
                    text = page.extract_text()
                    if text and text.strip():
                        sheet_name = f'Page {page_idx + 1}'
                        ws = wb.create_sheet(title=sheet_name)
                        lines = text.split('\n')
                        for row_idx, line in enumerate(lines):
                            if not line.strip():
                                continue
                            # Split by multiple spaces to detect columns
                            import re
                            parts = re.split(r'  +', line.strip())
                            for col_idx, part in enumerate(parts):
                                cell = ws.cell(
                                    row=row_idx + 1,
                                    column=col_idx + 1,
                                    value=part.strip()
                                )
                                cell.border = thin_border
                                cell.font = cell_font
                                cell.alignment = cell_align
                                if row_idx == 0:
                                    cell.font = header_font
                                    cell.fill = header_fill
                                    cell.alignment = header_align
                                elif row_idx % 2 == 0:
                                    cell.fill = alt_fill

                        # Auto-fit
                        if ws.max_column:
                            for col_idx in range(1, (ws.max_column or 0) + 1):
                                max_len = 8
                                col_letter = get_column_letter(col_idx)
                                for row_idx in range(1, (ws.max_row or 0) + 1):
                                    val = ws.cell(row=row_idx, column=col_idx).value
                                    if val:
                                        max_len = max(max_len, min(len(str(val)) + 2, 60))
                                ws.column_dimensions[col_letter].width = max_len
                        ws.freeze_panes = 'A2'

        if len(wb.sheetnames) == 0:
            ws = wb.create_sheet(title='Sheet1')
            ws['A1'] = 'No data could be extracted from this PDF.'
            ws['A1'].font = Font(italic=True, color='94A3B8')

        wb.save(output_path)
        return output_path

    except Exception as e:
        raise Exception(f'Failed to convert PDF to Excel: {str(e)}')


# ═══════════════════════════════════════════════════════════════
# 9. MERGE PDFs
# ═══════════════════════════════════════════════════════════════
def merge_pdfs(input_paths, original_name):
    """Merge multiple PDF files into a single PDF."""
    import fitz

    _, output_dir = ensure_media_dirs()
    base_name = Path(original_name).stem
    output_path = os.path.join(output_dir, f"{base_name}_merged.pdf")

    merged = fitz.open()
    for path in input_paths:
        pdf = fitz.open(path)
        merged.insert_pdf(pdf)
        pdf.close()

    merged.save(output_path)
    merged.close()
    return output_path


# ═══════════════════════════════════════════════════════════════
# 10. SPLIT PDF
# ═══════════════════════════════════════════════════════════════
def split_pdf(input_path, original_name, split_mode='each', page_ranges=None):
    """Split a PDF into individual pages or custom ranges. Returns a zip."""
    import fitz
    import zipfile

    _, output_dir = ensure_media_dirs()
    base_name = Path(original_name).stem
    zip_path = os.path.join(output_dir, f"{base_name}_split.zip")

    pdf = fitz.open(input_path)
    total_pages = len(pdf)

    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        if split_mode == 'ranges' and page_ranges:
            # Parse ranges like "1-3,5,7-9"
            for part in page_ranges.split(','):
                part = part.strip()
                if '-' in part:
                    start, end = part.split('-', 1)
                    start = max(1, int(start.strip()))
                    end = min(total_pages, int(end.strip()))
                else:
                    start = end = max(1, min(int(part.strip()), total_pages))

                out_pdf = fitz.open()
                for p in range(start - 1, end):
                    out_pdf.insert_pdf(pdf, from_page=p, to_page=p)

                fname = f"{base_name}_pages_{start}-{end}.pdf"
                tmp_path = os.path.join(output_dir, fname)
                out_pdf.save(tmp_path)
                out_pdf.close()
                zipf.write(tmp_path, fname)
                try:
                    os.remove(tmp_path)
                except OSError:
                    pass
        else:
            # Split each page
            for i in range(total_pages):
                out_pdf = fitz.open()
                out_pdf.insert_pdf(pdf, from_page=i, to_page=i)
                fname = f"{base_name}_page_{i + 1}.pdf"
                tmp_path = os.path.join(output_dir, fname)
                out_pdf.save(tmp_path)
                out_pdf.close()
                zipf.write(tmp_path, fname)
                try:
                    os.remove(tmp_path)
                except OSError:
                    pass

    pdf.close()
    return zip_path


# ═══════════════════════════════════════════════════════════════
# 11. COMPRESS PDF
# ═══════════════════════════════════════════════════════════════
def compress_pdf(input_path, original_name):
    """Compress a PDF file to reduce its size."""
    import fitz

    output_path = get_output_path(original_name, 'pdf')
    base_name = Path(original_name).stem
    output_path = os.path.join(
        os.path.dirname(output_path), f"{base_name}_compressed.pdf"
    )

    pdf = fitz.open(input_path)

    # Compress images on each page
    for page in pdf:
        images = page.get_images(full=True)
        for img_info in images:
            xref = img_info[0]
            try:
                base_image = pdf.extract_image(xref)
                if base_image and base_image.get("image"):
                    from PIL import Image
                    img_bytes = base_image["image"]
                    img = Image.open(io.BytesIO(img_bytes))

                    # Resize large images
                    max_dim = 1200
                    if img.width > max_dim or img.height > max_dim:
                        ratio = min(max_dim / img.width, max_dim / img.height)
                        new_size = (int(img.width * ratio), int(img.height * ratio))
                        img = img.resize(new_size, Image.LANCZOS)

                    # Convert to RGB if needed
                    if img.mode in ('RGBA', 'P', 'LA'):
                        img = img.convert('RGB')

                    buf = io.BytesIO()
                    img.save(buf, format='JPEG', quality=65, optimize=True)
                    buf.seek(0)

                    # Replace image in PDF
                    page.replace_image(xref, stream=buf.getvalue())
            except Exception:
                continue

    # Save with garbage collection and deflation
    pdf.save(
        output_path,
        garbage=4,
        deflate=True,
        deflate_images=True,
        deflate_fonts=True,
        clean=True,
    )
    pdf.close()
    return output_path


# ═══════════════════════════════════════════════════════════════
# 12. REMOVE PAGES FROM PDF
# ═══════════════════════════════════════════════════════════════
def remove_pdf_pages(input_path, original_name, pages_to_remove):
    """Remove specified pages from a PDF.

    pages_to_remove: comma-separated string like '1,3,5-7'
    """
    import fitz

    output_path = get_output_path(original_name, 'pdf')
    base_name = Path(original_name).stem
    output_path = os.path.join(
        os.path.dirname(output_path), f"{base_name}_trimmed.pdf"
    )

    pdf = fitz.open(input_path)
    total_pages = len(pdf)

    # Parse pages to remove (1-indexed input → 0-indexed)
    remove_set = set()
    for part in pages_to_remove.split(','):
        part = part.strip()
        if '-' in part:
            start, end = part.split('-', 1)
            for p in range(int(start.strip()), int(end.strip()) + 1):
                if 1 <= p <= total_pages:
                    remove_set.add(p - 1)
        else:
            p = int(part.strip())
            if 1 <= p <= total_pages:
                remove_set.add(p - 1)

    if len(remove_set) >= total_pages:
        raise Exception("Cannot remove all pages from the PDF.")

    # Build new PDF with remaining pages
    new_pdf = fitz.open()
    for i in range(total_pages):
        if i not in remove_set:
            new_pdf.insert_pdf(pdf, from_page=i, to_page=i)

    new_pdf.save(output_path)
    new_pdf.close()
    pdf.close()
    return output_path

