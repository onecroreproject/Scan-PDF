"""
Views for the file converter application.
"""
import os
import mimetypes
from django.shortcuts import render
from django.http import JsonResponse, FileResponse, Http404
from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.http import require_POST

from .forms import FileUploadForm
from .utils import (
    save_uploaded_file,
    convert_word_to_pdf,
    convert_pptx_to_pdf,
    convert_excel_to_pdf,
    convert_html_to_pdf,
    convert_pdf_to_image,
    convert_pdf_to_word,
    convert_pdf_to_pptx,
    convert_pdf_to_excel,
    merge_pdfs,
    split_pdf,
    compress_pdf,
    remove_pdf_pages,
    extract_pdf_pages,
    organize_pdf,
    repair_pdf,
    ocr_pdf,
    rotate_pdf,
    add_watermark,
    remove_watermark,
    crop_pdf,
)


# ─── Tool Configuration ────────────────────────────────────────
TOOLS = {
    'word-to-pdf': {
        'title': 'Word to PDF',
        'description': 'Convert Microsoft Word documents (.docx) to professional PDF files instantly.',
        'icon': 'file-text',
        'accept': '.docx',
        'allowed_extensions': ['.docx'],
        'converter': convert_word_to_pdf,
        'color': '#2b6cb0',
        'gradient': 'from-blue-500 to-blue-700',
        'category': 'convert',
    },
    'pptx-to-pdf': {
        'title': 'PowerPoint to PDF',
        'description': 'Transform PowerPoint presentations (.pptx) into shareable PDF documents.',
        'icon': 'presentation',
        'accept': '.pptx',
        'allowed_extensions': ['.pptx'],
        'converter': convert_pptx_to_pdf,
        'color': '#c05621',
        'gradient': 'from-orange-500 to-red-500',
        'category': 'convert',
    },
    'excel-to-pdf': {
        'title': 'Excel to PDF',
        'description': 'Convert Excel spreadsheets (.xlsx) to clean, formatted PDF files.',
        'icon': 'table',
        'accept': '.xlsx',
        'allowed_extensions': ['.xlsx'],
        'converter': convert_excel_to_pdf,
        'color': '#276749',
        'gradient': 'from-green-500 to-emerald-700',
        'category': 'convert',
    },
    'html-to-pdf': {
        'title': 'HTML to PDF',
        'description': 'Convert HTML files to pixel-perfect PDF documents with styling preserved.',
        'icon': 'code',
        'accept': '.html,.htm',
        'allowed_extensions': ['.html', '.htm'],
        'converter': convert_html_to_pdf,
        'color': '#6b46c1',
        'gradient': 'from-purple-500 to-indigo-700',
        'category': 'convert',
    },
    'pdf-to-image': {
        'title': 'PDF to Image',
        'description': 'Convert PDF pages to high-quality PNG or JPG images effortlessly.',
        'icon': 'image',
        'accept': '.pdf',
        'allowed_extensions': ['.pdf'],
        'converter': convert_pdf_to_image,
        'color': '#b83280',
        'gradient': 'from-pink-500 to-rose-600',
        'category': 'convert',
    },
    'pdf-to-word': {
        'title': 'PDF to Word',
        'description': 'Convert PDF files to editable Word documents (.docx) with accurate formatting.',
        'icon': 'file-type',
        'accept': '.pdf',
        'allowed_extensions': ['.pdf'],
        'converter': convert_pdf_to_word,
        'color': '#0d9488',
        'gradient': 'from-teal-500 to-teal-700',
        'category': 'convert',
    },
    'pdf-to-pptx': {
        'title': 'PDF to PowerPoint',
        'description': 'Transform PDF files into editable PowerPoint presentations (.pptx).',
        'icon': 'monitor-play',
        'accept': '.pdf',
        'allowed_extensions': ['.pdf'],
        'converter': convert_pdf_to_pptx,
        'color': '#d97706',
        'gradient': 'from-amber-500 to-amber-700',
        'category': 'convert',
    },
    'pdf-to-excel': {
        'title': 'PDF to Excel',
        'description': 'Extract tables from PDF files into editable Excel workbooks (.xlsx).',
        'icon': 'sheet',
        'accept': '.pdf',
        'allowed_extensions': ['.pdf'],
        'converter': convert_pdf_to_excel,
        'color': '#0891b2',
        'gradient': 'from-cyan-500 to-cyan-700',
        'category': 'convert',
    },
    'merge-pdf': {
        'title': 'Merge PDF',
        'description': 'Combine multiple PDF files into a single document in your desired order.',
        'icon': 'combine',
        'accept': '.pdf',
        'allowed_extensions': ['.pdf'],
        'converter': None,
        'color': '#7c3aed',
        'gradient': 'from-violet-500 to-violet-700',
        'category': 'pdf-tools',
        'multi_file': True,
    },
    'split-pdf': {
        'title': 'Split PDF',
        'description': 'Split a PDF into individual pages or custom page ranges instantly.',
        'icon': 'split',
        'accept': '.pdf',
        'allowed_extensions': ['.pdf'],
        'converter': None,
        'color': '#059669',
        'gradient': 'from-emerald-500 to-emerald-700',
        'category': 'pdf-tools',
    },
    'compress-pdf': {
        'title': 'Compress PDF',
        'description': 'Reduce your PDF file size while maintaining visual quality.',
        'icon': 'archive',
        'accept': '.pdf',
        'allowed_extensions': ['.pdf'],
        'converter': compress_pdf,
        'color': '#dc2626',
        'gradient': 'from-red-500 to-red-700',
        'category': 'pdf-tools',
    },
    'remove-pages': {
        'title': 'Remove Pages',
        'description': 'Delete specific pages from your PDF document easily.',
        'icon': 'file-minus',
        'accept': '.pdf',
        'allowed_extensions': ['.pdf'],
        'converter': None,
        'color': '#ea580c',
        'gradient': 'from-orange-500 to-orange-700',
        'category': 'pdf-tools',
    },
    'extract-pages': {
        'title': 'Extract Pages',
        'description': 'Pull specific pages out of a PDF into a new file.',
        'icon': 'file-output',
        'accept': '.pdf',
        'allowed_extensions': ['.pdf'],
        'converter': None,
        'color': '#0891b2',
        'gradient': 'from-cyan-500 to-cyan-700',
        'category': 'pdf-tools',
    },
    'organize-pdf': {
        'title': 'Organize PDF',
        'description': 'Reorder and rearrange the pages of your PDF effortlessly.',
        'icon': 'arrow-up-down',
        'accept': '.pdf',
        'allowed_extensions': ['.pdf'],
        'converter': None,
        'color': '#7c3aed',
        'gradient': 'from-violet-500 to-violet-700',
        'category': 'pdf-tools',
    },
    'repair-pdf': {
        'title': 'Repair PDF',
        'description': 'Fix corrupted or broken PDF files and recover content.',
        'icon': 'wrench',
        'accept': '.pdf',
        'allowed_extensions': ['.pdf'],
        'converter': repair_pdf,
        'color': '#b91c1c',
        'gradient': 'from-rose-500 to-rose-700',
        'category': 'pdf-tools',
    },
    'ocr-pdf': {
        'title': 'OCR PDF',
        'description': 'Convert scanned PDFs into searchable, selectable text.',
        'icon': 'scan-text',
        'accept': '.pdf',
        'allowed_extensions': ['.pdf'],
        'converter': ocr_pdf,
        'color': '#0d9488',
        'gradient': 'from-teal-500 to-teal-700',
        'category': 'pdf-tools',
    },
    'rotate-pdf': {
        'title': 'Rotate PDF',
        'description': 'Rotate all or specific pages of your PDF by 90°, 180°, or 270°.',
        'icon': 'rotate-cw',
        'accept': '.pdf',
        'allowed_extensions': ['.pdf'],
        'converter': None,
        'color': '#6366f1',
        'gradient': 'from-indigo-500 to-indigo-700',
        'category': 'pdf-tools',
    },
    'add-watermark': {
        'title': 'Add Watermark',
        'description': 'Overlay a custom text watermark on every page of your PDF.',
        'icon': 'stamp',
        'accept': '.pdf',
        'allowed_extensions': ['.pdf'],
        'converter': None,
        'color': '#0ea5e9',
        'gradient': 'from-sky-500 to-sky-700',
        'category': 'pdf-tools',
    },
    'remove-watermark': {
        'title': 'Remove Watermark',
        'description': 'Attempt to detect and remove watermarks from your PDF.',
        'icon': 'eraser',
        'accept': '.pdf',
        'allowed_extensions': ['.pdf'],
        'converter': remove_watermark,
        'color': '#f43f5e',
        'gradient': 'from-rose-500 to-rose-700',
        'category': 'pdf-tools',
    },
    'crop-pdf': {
        'title': 'Crop PDF',
        'description': 'Crop whitespace or set custom margins to resize your PDF pages.',
        'icon': 'crop',
        'accept': '.pdf',
        'allowed_extensions': ['.pdf'],
        'converter': None,
        'color': '#84cc16',
        'gradient': 'from-lime-500 to-lime-700',
        'category': 'pdf-tools',
    },
}


def home(request):
    """Render the home page with all available tools."""
    context = {
        'tools': TOOLS,
        'page_title': 'ScanPDF — Scan Smart. Share Fast.',
    }
    return render(request, 'converter/home.html', context)


def convert_page(request, tool_slug):
    """Render the conversion page for a specific tool."""
    if tool_slug not in TOOLS:
        raise Http404("Tool not found")

    tool = TOOLS[tool_slug]
    form = FileUploadForm()

    # Determine which template to use
    if tool_slug == 'merge-pdf':
        template = 'converter/merge.html'
    elif tool_slug == 'split-pdf':
        template = 'converter/split.html'
    elif tool_slug == 'remove-pages':
        template = 'converter/remove_pages.html'
    elif tool_slug == 'extract-pages':
        template = 'converter/extract_pages.html'
    elif tool_slug == 'organize-pdf':
        template = 'converter/organize_pdf.html'
    elif tool_slug == 'rotate-pdf':
        template = 'converter/rotate_pdf.html'
    elif tool_slug == 'add-watermark':
        template = 'converter/add_watermark.html'
    elif tool_slug == 'crop-pdf':
        template = 'converter/crop_pdf.html'
    else:
        template = 'converter/convert.html'

    context = {
        'tool': tool,
        'tool_slug': tool_slug,
        'form': form,
        'page_title': f'{tool["title"]} — ScanPDF',
    }
    return render(request, template, context)


@require_POST
def convert_file(request, tool_slug):
    """Handle file conversion via AJAX request."""
    if tool_slug not in TOOLS:
        return JsonResponse({'error': 'Invalid tool selected.'}, status=400)

    tool = TOOLS[tool_slug]

    # ── Merge PDF: multiple files ──
    if tool_slug == 'merge-pdf':
        files = request.FILES.getlist('files')
        if not files or len(files) < 2:
            return JsonResponse({'error': 'Please upload at least 2 PDF files to merge.'}, status=400)

        try:
            input_paths = []
            for f in files:
                ext = os.path.splitext(f.name)[1].lower()
                if ext != '.pdf':
                    return JsonResponse({'error': f'Invalid file "{f.name}". Only PDF files are allowed.'}, status=400)
                input_paths.append(save_uploaded_file(f))

            output_path = merge_pdfs(input_paths, files[0].name)

            for p in input_paths:
                try:
                    os.remove(p)
                except OSError:
                    pass

            content_type = 'application/pdf'
            output_filename = os.path.basename(output_path)
            response = FileResponse(open(output_path, 'rb'), content_type=content_type)
            response['Content-Disposition'] = f'attachment; filename="{output_filename}"'
            return response
        except Exception as e:
            return JsonResponse({'error': f'Merge failed: {str(e)}'}, status=500)

    # ── Split PDF ──
    if tool_slug == 'split-pdf':
        if 'file' not in request.FILES:
            return JsonResponse({'error': 'No file was uploaded.'}, status=400)

        uploaded_file = request.FILES['file']
        split_mode = request.POST.get('split_mode', 'each')
        page_ranges = request.POST.get('page_ranges', '')

        try:
            input_path = save_uploaded_file(uploaded_file)
            output_path = split_pdf(input_path, uploaded_file.name, split_mode, page_ranges)

            try:
                os.remove(input_path)
            except OSError:
                pass

            response = FileResponse(open(output_path, 'rb'), content_type='application/zip')
            output_filename = os.path.basename(output_path)
            response['Content-Disposition'] = f'attachment; filename="{output_filename}"'
            return response
        except Exception as e:
            return JsonResponse({'error': f'Split failed: {str(e)}'}, status=500)

    # ── Remove Pages ──
    if tool_slug == 'remove-pages':
        if 'file' not in request.FILES:
            return JsonResponse({'error': 'No file was uploaded.'}, status=400)

        uploaded_file = request.FILES['file']
        pages_to_remove = request.POST.get('pages_to_remove', '')

        if not pages_to_remove.strip():
            return JsonResponse({'error': 'Please specify which pages to remove.'}, status=400)

        try:
            input_path = save_uploaded_file(uploaded_file)
            output_path = remove_pdf_pages(input_path, uploaded_file.name, pages_to_remove)

            try:
                os.remove(input_path)
            except OSError:
                pass

            response = FileResponse(open(output_path, 'rb'), content_type='application/pdf')
            output_filename = os.path.basename(output_path)
            response['Content-Disposition'] = f'attachment; filename="{output_filename}"'
            return response
        except Exception as e:
            return JsonResponse({'error': f'Remove pages failed: {str(e)}'}, status=500)

    # ── Extract Pages ──
    if tool_slug == 'extract-pages':
        if 'file' not in request.FILES:
            return JsonResponse({'error': 'No file was uploaded.'}, status=400)

        uploaded_file = request.FILES['file']
        pages_to_extract = request.POST.get('pages_to_extract', '')

        if not pages_to_extract.strip():
            return JsonResponse({'error': 'Please specify which pages to extract.'}, status=400)

        try:
            input_path = save_uploaded_file(uploaded_file)
            output_path = extract_pdf_pages(input_path, uploaded_file.name, pages_to_extract)

            try:
                os.remove(input_path)
            except OSError:
                pass

            response = FileResponse(open(output_path, 'rb'), content_type='application/pdf')
            output_filename = os.path.basename(output_path)
            response['Content-Disposition'] = f'attachment; filename="{output_filename}"'
            return response
        except Exception as e:
            return JsonResponse({'error': f'Extract pages failed: {str(e)}'}, status=500)

    # ── Organize PDF ──
    if tool_slug == 'organize-pdf':
        if 'file' not in request.FILES:
            return JsonResponse({'error': 'No file was uploaded.'}, status=400)

        uploaded_file = request.FILES['file']
        page_order = request.POST.get('page_order', '')

        if not page_order.strip():
            return JsonResponse({'error': 'Please specify the desired page order.'}, status=400)

        try:
            input_path = save_uploaded_file(uploaded_file)
            output_path = organize_pdf(input_path, uploaded_file.name, page_order)

            try:
                os.remove(input_path)
            except OSError:
                pass

            response = FileResponse(open(output_path, 'rb'), content_type='application/pdf')
            output_filename = os.path.basename(output_path)
            response['Content-Disposition'] = f'attachment; filename="{output_filename}"'
            return response
        except Exception as e:
            return JsonResponse({'error': f'Organize PDF failed: {str(e)}'}, status=500)

    # ── Rotate PDF ──
    if tool_slug == 'rotate-pdf':
        if 'file' not in request.FILES:
            return JsonResponse({'error': 'No file was uploaded.'}, status=400)

        uploaded_file = request.FILES['file']
        rotation_angle = request.POST.get('rotation_angle', '90')
        page_selection = request.POST.get('page_selection', 'all')

        try:
            input_path = save_uploaded_file(uploaded_file)
            output_path = rotate_pdf(input_path, uploaded_file.name, rotation_angle, page_selection)

            try:
                os.remove(input_path)
            except OSError:
                pass

            response = FileResponse(open(output_path, 'rb'), content_type='application/pdf')
            output_filename = os.path.basename(output_path)
            response['Content-Disposition'] = f'attachment; filename="{output_filename}"'
            return response
        except Exception as e:
            return JsonResponse({'error': f'Rotate PDF failed: {str(e)}'}, status=500)

    # ── Add Watermark ──
    if tool_slug == 'add-watermark':
        if 'file' not in request.FILES:
            return JsonResponse({'error': 'No file was uploaded.'}, status=400)

        uploaded_file = request.FILES['file']
        watermark_text = request.POST.get('watermark_text', 'CONFIDENTIAL')
        opacity = request.POST.get('opacity', '0.15')
        font_size = request.POST.get('font_size', '60')
        rotation = request.POST.get('rotation', '45')
        color = request.POST.get('color', '#888888')

        if not watermark_text.strip():
            return JsonResponse({'error': 'Please enter watermark text.'}, status=400)

        try:
            input_path = save_uploaded_file(uploaded_file)
            output_path = add_watermark(
                input_path, uploaded_file.name,
                watermark_text=watermark_text,
                opacity=opacity,
                font_size=font_size,
                rotation=rotation,
                color=color,
            )

            try:
                os.remove(input_path)
            except OSError:
                pass

            response = FileResponse(open(output_path, 'rb'), content_type='application/pdf')
            output_filename = os.path.basename(output_path)
            response['Content-Disposition'] = f'attachment; filename="{output_filename}"'
            return response
        except Exception as e:
            return JsonResponse({'error': f'Add watermark failed: {str(e)}'}, status=500)

    # ── Crop PDF ──
    if tool_slug == 'crop-pdf':
        if 'file' not in request.FILES:
            return JsonResponse({'error': 'No file was uploaded.'}, status=400)

        uploaded_file = request.FILES['file']
        crop_mode = request.POST.get('crop_mode', 'auto')
        crop_top = request.POST.get('crop_top', '0')
        crop_bottom = request.POST.get('crop_bottom', '0')
        crop_left = request.POST.get('crop_left', '0')
        crop_right = request.POST.get('crop_right', '0')

        try:
            input_path = save_uploaded_file(uploaded_file)
            output_path = crop_pdf(
                input_path, uploaded_file.name,
                crop_mode=crop_mode,
                top=crop_top,
                bottom=crop_bottom,
                left=crop_left,
                right=crop_right,
            )

            try:
                os.remove(input_path)
            except OSError:
                pass

            response = FileResponse(open(output_path, 'rb'), content_type='application/pdf')
            output_filename = os.path.basename(output_path)
            response['Content-Disposition'] = f'attachment; filename="{output_filename}"'
            return response
        except Exception as e:
            return JsonResponse({'error': f'Crop PDF failed: {str(e)}'}, status=500)

    # ── Standard single-file conversion ──
    if 'file' not in request.FILES:
        return JsonResponse({'error': 'No file was uploaded. Please select a file.'}, status=400)

    uploaded_file = request.FILES['file']

    file_ext = os.path.splitext(uploaded_file.name)[1].lower()
    if file_ext not in tool['allowed_extensions']:
        allowed = ', '.join(tool['allowed_extensions'])
        return JsonResponse({
            'error': f'Invalid file type "{file_ext}". Allowed types: {allowed}'
        }, status=400)

    if uploaded_file.size > 52428800:
        return JsonResponse({
            'error': 'File size exceeds the 50MB limit. Please upload a smaller file.'
        }, status=400)

    try:
        input_path = save_uploaded_file(uploaded_file)
        output_path = tool['converter'](input_path, uploaded_file.name)

        try:
            os.remove(input_path)
        except OSError:
            pass

        content_type, _ = mimetypes.guess_type(output_path)
        if content_type is None:
            content_type = 'application/octet-stream'

        output_filename = os.path.basename(output_path)
        response = FileResponse(
            open(output_path, 'rb'),
            content_type=content_type,
        )
        response['Content-Disposition'] = f'attachment; filename="{output_filename}"'
        return response

    except Exception as e:
        return JsonResponse({
            'error': f'Conversion failed: {str(e)}'
        }, status=500)
