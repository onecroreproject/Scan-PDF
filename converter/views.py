"""
Views for the file converter application.
"""
import os
import mimetypes
from pathlib import Path
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
    edit_pdf,
    convert_pdf_to_html_via_word,
    convert_html_to_pdf_from_string,
    unlock_pdf,
    protect_pdf,
    png_to_jpg,
    jpg_to_png,
    html_to_image,
    resize_image,
    scale_image,
    rotate_image,
    add_image_watermark,
    compress_image,
    crop_image,
    remove_background,
    balance_chemical_equation,
    generate_qr_code,
    generate_meme,
    generate_password,
    generate_story,
    generate_names,
    get_video_info,
    download_video,
    run_speed_test,
    convert_images_to_pdf,
    convert_pdf_to_pdfa,
    sign_pdf,
    redact_pdf,
)


class FileCleanupResponse(FileResponse):
    """
    A specialization of FileResponse that deletes the underlying file on disk
    once the response has been closed.
    """
    def __init__(self, file_path, *args, **kwargs):
        self._temp_file_path = file_path
        # We must open the file first to pass it to the parent constructor
        file_handle = open(file_path, 'rb')
        super().__init__(file_handle, *args, **kwargs)

    def close(self):
        super().close()
        # After the response is closed (stream finished), delete the file
        if self._temp_file_path and os.path.exists(self._temp_file_path):
            try:
                os.remove(self._temp_file_path)
            except OSError:
                pass


def create_cleanup_response(file_path, content_type=None, filename=None):
    """Helper to create a cleanup response with proper headers and formatted filenames."""
    if not content_type:
        import mimetypes
        content_type, _ = mimetypes.guess_type(file_path)
        content_type = content_type or 'application/octet-stream'
    
    # Choose the starting point for formatting: passed filename or the disk name
    raw_name = filename or os.path.basename(file_path)
    
    # Use the formatting logic from utils
    from .utils import format_download_name
    final_filename = format_download_name(raw_name)

    response = FileCleanupResponse(file_path, content_type=content_type)
    response['Content-Disposition'] = f'attachment; filename="{final_filename}"'
    return response



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
        'description': 'Convert any webpage URL or HTML file to a pixel-perfect PDF document.',
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
        'title': 'OCR to PDF',
        'description': 'Convert scanned PDFs, Word documents, and images into searchable, selectable PDF documents or extract their text directly.',
        'icon': 'scan-text',
        'accept': '.pdf,.jpg,.jpeg,.png,.docx',
        'allowed_extensions': ['.pdf', '.jpg', '.jpeg', '.png', '.docx'],
        'converter': ocr_pdf,
        'color': '#0d9488',
        'gradient': 'from-teal-500 to-teal-700',
        'category': 'pdf-tools',
        'multi_file': True,
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
    'edit-pdf': {
        'title': 'Edit PDF',
        'description': 'Add text annotations and notes to your PDF pages.',
        'icon': 'pencil',
        'accept': '.pdf',
        'allowed_extensions': ['.pdf'],
        'converter': None,
        'color': '#8b5cf6',
        'gradient': 'from-violet-500 to-purple-700',
        'category': 'pdf-tools',
    },
    'unlock-pdf': {
        'title': 'Unlock PDF',
        'description': 'Remove password protection from your secured PDF files.',
        'icon': 'lock-open',
        'accept': '.pdf',
        'allowed_extensions': ['.pdf'],
        'converter': None,
        'color': '#10b981',
        'gradient': 'from-emerald-500 to-emerald-700',
        'category': 'pdf-tools',
    },
    'protect-pdf': {
        'title': 'Protect PDF',
        'description': 'Encrypt your PDF with a password to restrict access.',
        'icon': 'shield-check',
        'accept': '.pdf',
        'allowed_extensions': ['.pdf'],
        'converter': None,
        'color': '#ef4444',
        'gradient': 'from-red-500 to-red-700',
        'category': 'pdf-tools',
    },
    'png-to-jpg': {
        'title': 'PNG to JPG',
        'description': 'Convert PNG images to high-quality JPEG format instantly.',
        'icon': 'image',
        'accept': '.png',
        'allowed_extensions': ['.png'],
        'converter': png_to_jpg,
        'color': '#2b6cb0',
        'gradient': 'from-blue-500 to-blue-700',
        'category': 'convert',
    },
    'jpg-to-png': {
        'title': 'JPG to PNG',
        'description': 'Convert JPEG images to PNG format with lossless quality.',
        'icon': 'image',
        'accept': '.jpg,.jpeg',
        'allowed_extensions': ['.jpg', '.jpeg'],
        'converter': jpg_to_png,
        'color': '#276749',
        'gradient': 'from-green-500 to-emerald-700',
        'category': 'convert',
    },
    'html-to-image': {
        'title': 'HTML to Image',
        'description': 'Capture a pixel-perfect image of your HTML files.',
        'icon': 'file-code',
        'accept': '.html,.htm',
        'allowed_extensions': ['.html', '.htm'],
        'converter': html_to_image,
        'color': '#c05621',
        'gradient': 'from-orange-500 to-red-500',
        'category': 'convert',
    },
    'resize-image': {
        'title': 'Resize Image',
        'description': 'Set an exact width and height for your JPG images.',
        'icon': 'move',
        'accept': '.jpg,.jpeg,.png',
        'allowed_extensions': ['.jpg', '.jpeg', '.png'],
        'converter': None,
        'color': '#0d9488',
        'gradient': 'from-teal-500 to-teal-700',
        'category': 'image-tools',
    },
    'scale-image': {
        'title': 'Scale Image',
        'description': 'Scale your image up or down by a percentage.',
        'icon': 'maximize-2',
        'accept': '.jpg,.jpeg,.png',
        'allowed_extensions': ['.jpg', '.jpeg', '.png'],
        'converter': None,
        'color': '#7c3aed',
        'gradient': 'from-violet-500 to-violet-700',
        'category': 'image-tools',
    },
    'rotate-image': {
        'title': 'Rotate Image',
        'description': 'Rotate your image by any angle with one click.',
        'icon': 'rotate-cw',
        'accept': '.jpg,.jpeg,.png',
        'allowed_extensions': ['.jpg', '.jpeg', '.png'],
        'converter': None,
        'color': '#6366f1',
        'gradient': 'from-indigo-500 to-indigo-700',
        'category': 'image-tools',
    },
    'add-image-watermark': {
        'title': 'Add Watermark',
        'description': 'Overlay a custom text watermark on your images.',
        'icon': 'stamp',
        'accept': '.jpg,.jpeg,.png',
        'allowed_extensions': ['.jpg', '.jpeg', '.png'],
        'converter': None,
        'color': '#0ea5e9',
        'gradient': 'from-sky-500 to-sky-700',
        'category': 'image-tools',
    },
    'compress-image': {
        'title': 'Compress Image',
        'description': 'Reduce your image file size while keeping great quality.',
        'icon': 'archive',
        'accept': '.jpg,.jpeg,.png',
        'allowed_extensions': ['.jpg', '.jpeg', '.png'],
        'converter': None,
        'color': '#dc2626',
        'gradient': 'from-red-500 to-red-700',
        'category': 'image-tools',
    },
    'crop-image': {
        'title': 'Crop Image',
        'description': 'Crop your image to a precise rectangle selection.',
        'icon': 'crop',
        'accept': '.jpg,.jpeg,.png',
        'allowed_extensions': ['.jpg', '.jpeg', '.png'],
        'converter': None,
        'color': '#84cc16',
        'gradient': 'from-lime-500 to-lime-700',
        'category': 'image-tools',
    },
    'remove-bg': {
        'title': 'Remove Background',
        'description': 'Instantly remove the background from any image using AI.',
        'icon': 'eraser',
        'accept': '.jpg,.jpeg,.png',
        'allowed_extensions': ['.jpg', '.jpeg', '.png'],
        'converter': remove_background,
        'color': '#f43f5e',
        'gradient': 'from-rose-500 to-rose-700',
        'category': 'image-tools',
    },
    'chemical-balancer': {
        'title': 'Chemical Balance',
        'description': 'Balance chemical equations instantly with stoichiometry.',
        'icon': 'beaker',
        'accept': None,
        'allowed_extensions': [],
        'converter': None,
        'color': '#8b5cf6',
        'gradient': 'from-violet-500 to-purple-600',
        'category': 'generate',
    },
    'password-generator': {
        'title': 'Password Generator',
        'description': 'Create secure, random passwords for your accounts.',
        'icon': 'key',
        'accept': None,
        'allowed_extensions': [],
        'converter': None,
        'color': '#059669',
        'gradient': 'from-emerald-500 to-teal-600',
        'category': 'generate',
    },
    'unit-converter': {
        'title': 'Unit Converter',
        'description': 'Convert between length, weight, temp, and more.',
        'icon': 'ruler',
        'accept': None,
        'allowed_extensions': [],
        'converter': None,
        'color': '#3b82f6',
        'gradient': 'from-blue-500 to-indigo-600',
        'category': 'other',
    },
    'speed-test': {
        'title': 'Speed Test',
        'description': 'Check your internet connection speed in seconds.',
        'icon': 'zap',
        'accept': None,
        'allowed_extensions': [],
        'converter': None,
        'color': '#f59e0b',
        'gradient': 'from-amber-400 to-orange-500',
        'category': 'other',
    },
    'instagram-downloader': {
        'title': 'IG Reels Downloader',
        'description': 'Save Instagram Reels and videos directly to your device.',
        'icon': 'instagram',
        'accept': None,
        'allowed_extensions': [],
        'converter': None,
        'color': '#d946ef',
        'gradient': 'from-fuchsia-500 to-pink-600',
        'category': 'download',
    },
    'youtube-downloader': {
        'title': 'YouTube Downloader',
        'description': 'Download YouTube videos in high quality.',
        'icon': 'video',
        'accept': None,
        'allowed_extensions': [],
        'converter': None,
        'color': '#ef4444',
        'gradient': 'from-red-500 to-rose-600',
        'category': 'download',
    },
    'qrcode-generator': {
        'title': 'QR Code Generator',
        'description': 'Generate custom QR codes for links, text, or Wi-Fi.',
        'icon': 'qr-code',
        'accept': None,
        'allowed_extensions': [],
        'converter': None,
        'color': '#111827',
        'gradient': 'from-gray-700 to-black',
        'category': 'generate',
    },
    'meme-generator': {
        'title': 'Meme Generator',
        'description': 'Create funny memes by adding text to your images.',
        'icon': 'laugh',
        'accept': '.jpg,.jpeg,.png',
        'allowed_extensions': ['.jpg', '.jpeg', '.png'],
        'converter': None,
        'color': '#facc15',
        'gradient': 'from-yellow-400 to-yellow-600',
        'category': 'generate',
    },
    'name-generator': {
        'title': 'Name Generator',
        'description': 'Generate random names for people, places, or companies.',
        'icon': 'user-plus',
        'accept': None,
        'allowed_extensions': [],
        'converter': None,
        'color': '#2dd4bf',
        'gradient': 'from-teal-400 to-cyan-500',
        'category': 'generate',
    },
    'image-to-video': {
        'title': 'AI Video Gen',
        'description': 'AI Video generation is coming soon! Stay tuned for cinematic video creation.',
        'icon': 'monitor-play',
        'accept': None,
        'allowed_extensions': [],
        'converter': None,
        'color': '#8b5cf6',
        'gradient': 'from-violet-500 to-purple-600',
        'category': 'ai-tools',
        'is_coming_soon': True,
    },
    'story-generator': {
        'title': 'AI Story Generator',
        'description': 'Generate creative stories from different genres using Gemini AI.',
        'icon': 'book-open',
        'accept': None,
        'allowed_extensions': [],
        'converter': generate_story,
        'color': '#4ade80',
        'gradient': 'from-green-400 to-emerald-500',
        'category': 'ai-tools',
    },
    'image-to-pdf': {
        'title': 'Image to PDF',
        'description': 'Convert one or more images (.jpg, .png) into a single PDF document.',
        'icon': 'file-up',
        'accept': '.jpg,.jpeg,.png',
        'allowed_extensions': ['.jpg', '.jpeg', '.png'],
        'converter': convert_images_to_pdf,
        'color': '#0ea5e9',
        'gradient': 'from-sky-500 to-indigo-600',
        'category': 'convert',
        'multi_file': True,
    },
    'pdf-to-pdfa': {
        'title': 'PDF to PDF/A',
        'description': 'Convert your PDF to PDF/A archival format for long-term preservation and compliance.',
        'icon': 'archive',
        'accept': '.pdf',
        'allowed_extensions': ['.pdf'],
        'converter': convert_pdf_to_pdfa,
        'color': '#0d9488',
        'gradient': 'from-teal-500 to-teal-700',
        'category': 'convert',
    },
    'sign-pdf': {
        'title': 'Sign PDF',
        'description': 'Draw, type, or upload a signature and place it on any page of your PDF.',
        'icon': 'pen-tool',
        'accept': '.pdf',
        'allowed_extensions': ['.pdf'],
        'converter': None,
        'color': '#059669',
        'gradient': 'from-emerald-500 to-emerald-700',
        'category': 'pdf-tools',
    },
    'redact-pdf': {
        'title': 'Redact PDF',
        'description': 'Permanently black out sensitive text and areas in your PDF documents.',
        'icon': 'eye-off',
        'accept': '.pdf',
        'allowed_extensions': ['.pdf'],
        'converter': None,
        'color': '#dc2626',
        'gradient': 'from-red-600 to-red-800',
        'category': 'pdf-tools',
    },
}


def home(request):
    """Render the home page with all available tools."""
    context = {
        'tools': TOOLS,
        'page_title': 'ScanPDF',
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
    elif tool_slug == 'edit-pdf':
        template = 'converter/edit_pdf.html'
    elif tool_slug == 'unlock-pdf':
        template = 'converter/unlock_pdf.html'
    elif tool_slug == 'protect-pdf':
        template = 'converter/protect_pdf.html'
    elif tool_slug == 'image-to-pdf':
        template = 'converter/image_to_pdf.html'
    elif tool_slug == 'ocr-pdf':
        template = 'converter/ocr.html'
    elif tool_slug == 'resize-image':
        template = 'converter/resize_image.html'
    elif tool_slug == 'scale-image':
        template = 'converter/scale_image.html'
    elif tool_slug == 'rotate-image':
        template = 'converter/rotate_image.html'
    elif tool_slug == 'add-image-watermark':
        template = 'converter/add_image_watermark.html'
    elif tool_slug == 'compress-image':
        template = 'converter/compress_image.html'
    elif tool_slug == 'crop-image':
        template = 'converter/crop_image.html'
    elif tool_slug == 'remove-bg':
        template = 'converter/remove_bg.html'
    elif tool_slug == 'chemical-balancer':
        template = 'converter/chemical_balancer.html'
    elif tool_slug == 'password-generator':
        template = 'converter/password_generator.html'
    elif tool_slug == 'unit-converter':
        template = 'converter/unit_converter.html'
    elif tool_slug == 'speed-test':
        template = 'converter/speed_test.html'
    elif tool_slug == 'instagram-downloader':
        template = 'converter/downloader.html'
    elif tool_slug == 'youtube-downloader':
        template = 'converter/downloader.html'
    elif tool_slug == 'qrcode-generator':
        template = 'converter/qrcode_generator.html'
    elif tool_slug == 'meme-generator':
        template = 'converter/meme_generator.html'
    elif tool_slug == 'story-generator':
        template = 'converter/story_generator.html'
    elif tool_slug == 'name-generator':
        template = 'converter/name_generator.html'
    elif tool_slug == 'image-to-video':
        template = 'converter/image_to_video.html'
    elif tool_slug == 'sign-pdf':
        template = 'converter/sign_pdf.html'
    elif tool_slug == 'redact-pdf':
        template = 'converter/redact_pdf.html'
    elif tool_slug == 'html-to-pdf':
        template = 'converter/html_to_pdf.html'
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

            return create_cleanup_response(output_path, content_type='application/pdf',
                                           filename=f"{Path(files[0].name).stem}_merged.pdf")
        except Exception as e:
            return JsonResponse({'error': f'Merge failed: {str(e)}'}, status=500)

    # ── HTML to PDF (URL or file) ──
    if tool_slug == 'html-to-pdf':
        url_input = request.POST.get('url', '').strip()
        uploaded_file = request.FILES.get('file')

        if not url_input and not uploaded_file:
            return JsonResponse({'error': 'Please provide a URL or upload an HTML file.'}, status=400)

        try:
            if url_input:
                # URL mode
                if not url_input.startswith(('http://', 'https://')):
                    url_input = 'https://' + url_input
                from urllib.parse import urlparse
                domain = urlparse(url_input).netloc or 'webpage'
                output_path = convert_html_to_pdf(None, f"{domain}.html", url=url_input)
            else:
                # File mode
                input_path = save_uploaded_file(uploaded_file)
                output_path = convert_html_to_pdf(input_path, uploaded_file.name)
                try:
                    os.remove(input_path)
                except OSError:
                    pass

            return create_cleanup_response(output_path, content_type='application/pdf')
        except Exception as e:
            return JsonResponse({'error': f'HTML to PDF failed: {str(e)}'}, status=500)

    # ── HTML to Image (URL or file) ──
    if tool_slug == 'html-to-image':
        url_input = request.POST.get('url', '').strip()
        uploaded_file = request.FILES.get('file')

        if not url_input and not uploaded_file:
            return JsonResponse({'error': 'Please provide a URL or upload an HTML file.'}, status=400)

        try:
            if url_input:
                # URL mode
                if not url_input.startswith(('http://', 'https://')):
                    url_input = 'https://' + url_input
                from urllib.parse import urlparse
                domain = urlparse(url_input).netloc or 'webpage'
                output_path = html_to_image(None, f"{domain}.png", url=url_input)
            else:
                # File mode
                input_path = save_uploaded_file(uploaded_file)
                output_path = html_to_image(input_path, uploaded_file.name)
                try:
                    os.remove(input_path)
                except OSError:
                    pass

            return create_cleanup_response(output_path, content_type='image/png')
        except Exception as e:
            return JsonResponse({'error': f'HTML to Image failed: {str(e)}'}, status=500)

    # ── Image to PDF: multiple files ──
    if tool_slug == 'image-to-pdf':
        files = request.FILES.getlist('files')
        if not files:
            files = [request.FILES.get('file')] if 'file' in request.FILES else []
        
        if not files:
            return JsonResponse({'error': 'Please upload at least one image.'}, status=400)

        try:
            input_paths = []
            for f in files:
                ext = os.path.splitext(f.name)[1].lower()
                if ext not in tool['allowed_extensions']:
                    return JsonResponse({'error': f'Invalid file "{f.name}". Only images (.jpg, .png) are allowed.'}, status=400)
                input_paths.append(save_uploaded_file(f))

            output_path = convert_images_to_pdf(input_paths, files[0].name)

            for p in input_paths:
                try:
                    os.remove(p)
                except OSError:
                    pass

            return create_cleanup_response(output_path, content_type='application/pdf')
        except Exception as e:
            return JsonResponse({'error': f'PDF creation failed: {str(e)}'}, status=500)

    # ── Image to GIF: multiple files ──
    if tool_slug == 'image-to-gif':
        files = request.FILES.getlist('files')
        if not files:
            # Fallback to single file if MultiValueDict is empty
            files = [request.FILES.get('file')] if 'file' in request.FILES else []

        if not files:
            return JsonResponse({'error': 'Please upload at least one image to create a GIF.'}, status=400)

        try:
            input_paths = []
            for f in files:
                ext = os.path.splitext(f.name)[1].lower()
                if ext not in tool['allowed_extensions']:
                    return JsonResponse({'error': f'Invalid file "{f.name}". Only images (.jpg, .png) are allowed.'}, status=400)
                input_paths.append(save_uploaded_file(f))

            output_path = image_to_gif(input_paths, files[0].name)

            for p in input_paths:
                try:
                    os.remove(p)
                except OSError:
                    pass

            return create_cleanup_response(output_path, content_type='image/gif')
        except Exception as e:
            return JsonResponse({'error': f'GIF creation failed: {str(e)}'}, status=500)

    # ── OCR to PDF: multiple files ──
    if tool_slug == 'ocr-pdf':
        files = request.FILES.getlist('files')
        if not files:
            files = [request.FILES.get('file')] if 'file' in request.FILES else []
        
        if not files:
            return JsonResponse({'error': 'Please upload at least one image, PDF, or Word file.'}, status=400)

        try:
            input_paths = []
            for f in files:
                ext = os.path.splitext(f.name)[1].lower()
                if ext not in tool['allowed_extensions']:
                    allowed = ', '.join(tool['allowed_extensions'])
                    return JsonResponse({'error': f'Invalid file "{f.name}". Allowed types: {allowed}'}, status=400)
                input_paths.append(save_uploaded_file(f))

            from .utils import extract_all_text
            extracted_text = extract_all_text(input_paths)

            for p in input_paths:
                try:
                    os.remove(p)
                except OSError:
                    pass

            return JsonResponse({'extracted_text': extracted_text})
        except Exception as e:
            return JsonResponse({'error': f'OCR failed: {str(e)}'}, status=500)

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

            return create_cleanup_response(output_path, content_type='application/zip')
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

            return create_cleanup_response(output_path, content_type='application/pdf')
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

            return create_cleanup_response(output_path, content_type='application/pdf')
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

            return create_cleanup_response(output_path, content_type='application/pdf')
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

            return create_cleanup_response(output_path, content_type='application/pdf')
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

            return create_cleanup_response(output_path, content_type='application/pdf')
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
        crop_x = request.POST.get('crop_x', '0')
        crop_y = request.POST.get('crop_y', '0')
        crop_w = request.POST.get('crop_w', '0')
        crop_h = request.POST.get('crop_h', '0')

        try:
            input_path = save_uploaded_file(uploaded_file)
            output_path = crop_pdf(
                input_path, uploaded_file.name,
                crop_mode=crop_mode,
                top=crop_top,
                bottom=crop_bottom,
                left=crop_left,
                right=crop_right,
                crop_x=crop_x,
                crop_y=crop_y,
                crop_w=crop_w,
                crop_h=crop_h,
            )

            try:
                os.remove(input_path)
            except OSError:
                pass

            return create_cleanup_response(output_path, content_type='application/pdf')
        except Exception as e:
            return JsonResponse({'error': f'Crop PDF failed: {str(e)}'}, status=500)

    # ── Edit PDF ──
    if tool_slug == 'edit-pdf':
        if 'file' not in request.FILES and 'html_content' not in request.POST:
            return JsonResponse({'error': 'No file or content provided.'}, status=400)

        html_content = request.POST.get('html_content')
        
        try:
            if html_content:
                # Case 2: User is downloading the edited content as PDF
                output_path = edit_pdf(None, "edited.pdf", html_content=html_content)
                return create_cleanup_response(output_path, content_type='application/pdf', filename="edited_document.pdf")
            else:
                # Case 1: Initial upload - convert PDF to editable HTML
                uploaded_file = request.FILES['file']
                input_path = save_uploaded_file(uploaded_file)
                html_data = convert_pdf_to_html_via_word(input_path)
                
                try:
                    os.remove(input_path)
                except OSError:
                    pass
                
                return JsonResponse({
                    'success': True,
                    'html': html_data,
                    'filename': uploaded_file.name
                })
        except Exception as e:
            return JsonResponse({'error': f'PDF Editor failed: {str(e)}'}, status=500)

    # ── Unlock PDF ──
    if tool_slug == 'unlock-pdf':
        if 'file' not in request.FILES:
            return JsonResponse({'error': 'No file was uploaded.'}, status=400)

        uploaded_file = request.FILES['file']
        password = request.POST.get('password', '')

        try:
            input_path = save_uploaded_file(uploaded_file)
            output_path = unlock_pdf(input_path, uploaded_file.name, password=password)

            try:
                os.remove(input_path)
            except OSError:
                pass

            return create_cleanup_response(output_path, content_type='application/pdf')
        except Exception as e:
            return JsonResponse({'error': f'Unlock PDF failed: {str(e)}'}, status=500)

    # ── Protect PDF ──
    if tool_slug == 'protect-pdf':
        if 'file' not in request.FILES:
            return JsonResponse({'error': 'No file was uploaded.'}, status=400)

        uploaded_file = request.FILES['file']
        user_password = request.POST.get('user_password', '')
        owner_password = request.POST.get('owner_password', '')

        if not user_password:
            return JsonResponse({'error': 'Please enter a password to protect this PDF.'}, status=400)

        try:
            input_path = save_uploaded_file(uploaded_file)
            output_path = protect_pdf(
                input_path, uploaded_file.name,
                user_password=user_password,
                owner_password=owner_password or user_password,
            )

            try:
                os.remove(input_path)
            except OSError:
                pass

            return create_cleanup_response(output_path, content_type='application/pdf')
        except Exception as e:
            return JsonResponse({'error': f'Protect PDF failed: {str(e)}'}, status=500)

    # ── Resize Image ──
    if tool_slug == 'resize-image':
        if 'file' not in request.FILES:
            return JsonResponse({'error': 'No file was uploaded.'}, status=400)

        uploaded_file = request.FILES['file']
        width = request.POST.get('width', '800')
        height = request.POST.get('height', '600')
        maintain_aspect = request.POST.get('maintain_aspect', 'true') == 'true'

        try:
            input_path = save_uploaded_file(uploaded_file)
            output_path = resize_image(
                input_path, uploaded_file.name,
                width=width, height=height,
                maintain_aspect=maintain_aspect,
            )
            try:
                os.remove(input_path)
            except OSError:
                pass

            return create_cleanup_response(output_path, content_type='image/jpeg')
        except Exception as e:
            return JsonResponse({'error': f'Resize failed: {str(e)}'}, status=500)

    # ── Scale Image ──
    if tool_slug == 'scale-image':
        if 'file' not in request.FILES:
            return JsonResponse({'error': 'No file was uploaded.'}, status=400)

        uploaded_file = request.FILES['file']
        scale_percent = request.POST.get('scale_percent', '50')

        try:
            input_path = save_uploaded_file(uploaded_file)
            output_path = scale_image(
                input_path, uploaded_file.name,
                scale_percent=scale_percent,
            )
            try:
                os.remove(input_path)
            except OSError:
                pass

            return create_cleanup_response(output_path, content_type='image/jpeg')
        except Exception as e:
            return JsonResponse({'error': f'Scale failed: {str(e)}'}, status=500)

    # ── Rotate Image ──
    if tool_slug == 'rotate-image':
        if 'file' not in request.FILES:
            return JsonResponse({'error': 'No file was uploaded.'}, status=400)

        uploaded_file = request.FILES['file']
        angle = request.POST.get('angle', '90')

        try:
            input_path = save_uploaded_file(uploaded_file)
            output_path = rotate_image(
                input_path, uploaded_file.name,
                angle=angle,
            )
            try:
                os.remove(input_path)
            except OSError:
                pass

            return create_cleanup_response(output_path, content_type='image/jpeg')
        except Exception as e:
            return JsonResponse({'error': f'Rotate failed: {str(e)}'}, status=500)

    # ── Add Image Watermark ──
    if tool_slug == 'add-image-watermark':
        if 'file' not in request.FILES:
            return JsonResponse({'error': 'No file was uploaded.'}, status=400)

        uploaded_file = request.FILES['file']
        watermark_text = request.POST.get('watermark_text', 'SAMPLE')
        opacity = request.POST.get('opacity', '0.3')
        font_size = request.POST.get('font_size', '40')
        color = request.POST.get('color', '#888888')

        if not watermark_text.strip():
            return JsonResponse({'error': 'Please enter watermark text.'}, status=400)

        try:
            input_path = save_uploaded_file(uploaded_file)
            output_path = add_image_watermark(
                input_path, uploaded_file.name,
                watermark_text=watermark_text,
                opacity=opacity, font_size=font_size, color=color,
            )
            try:
                os.remove(input_path)
            except OSError:
                pass

            return create_cleanup_response(output_path, content_type='image/jpeg')
        except Exception as e:
            return JsonResponse({'error': f'Add watermark failed: {str(e)}'}, status=500)

    # ── Compress Image ──
    if tool_slug == 'compress-image':
        if 'file' not in request.FILES:
            return JsonResponse({'error': 'No file was uploaded.'}, status=400)

        uploaded_file = request.FILES['file']
        quality = request.POST.get('quality', '60')

        try:
            input_path = save_uploaded_file(uploaded_file)
            output_path = compress_image(
                input_path, uploaded_file.name,
                quality=quality,
            )
            try:
                os.remove(input_path)
            except OSError:
                pass

            return create_cleanup_response(output_path, content_type='image/jpeg')
        except Exception as e:
            return JsonResponse({'error': f'Compress failed: {str(e)}'}, status=500)

    # ── Crop Image ──
    if tool_slug == 'crop-image':
        if 'file' not in request.FILES:
            return JsonResponse({'error': 'No file was uploaded.'}, status=400)

        uploaded_file = request.FILES['file']
        crop_x = request.POST.get('crop_x', '0')
        crop_y = request.POST.get('crop_y', '0')
        crop_width = request.POST.get('crop_width', '0')
        crop_height = request.POST.get('crop_height', '0')

        try:
            input_path = save_uploaded_file(uploaded_file)
            output_path = crop_image(
                input_path, uploaded_file.name,
                crop_x=crop_x, crop_y=crop_y,
                crop_width=crop_width, crop_height=crop_height,
            )
            try:
                os.remove(input_path)
            except OSError:
                pass

            return create_cleanup_response(output_path, content_type='image/jpeg')
        except Exception as e:
            return JsonResponse({'error': f'Crop failed: {str(e)}'}, status=500)

    # ── Remove Background ──
    if tool_slug == 'remove-bg':
        if 'file' not in request.FILES:
            return JsonResponse({'error': 'No file was uploaded.'}, status=400)

        uploaded_file = request.FILES['file']
        try:
            input_path = save_uploaded_file(uploaded_file)
            output_path = remove_background(input_path, uploaded_file.name)
            try:
                os.remove(input_path)
            except OSError:
                pass

            return create_cleanup_response(output_path, content_type='image/png')
        except Exception as e:
            return JsonResponse({'error': f'Background removal failed: {str(e)}'}, status=500)

    # ── Chemical Balance ──
    if tool_slug == 'chemical-balancer':
        equation = request.POST.get('equation', '')
        if not equation:
            return JsonResponse({'error': 'Please enter a chemical equation.'}, status=400)
        try:
            balanced = balance_chemical_equation(equation)
            return JsonResponse({'result': balanced})
        except Exception as e:
            return JsonResponse({'error': str(e)}, status=500)

    # ── Password Generator ──
    if tool_slug == 'password-generator':
        length = request.POST.get('length', 12)
        use_upper = request.POST.get('use_upper') == 'true'
        use_nums = request.POST.get('use_nums') == 'true'
        use_syms = request.POST.get('use_syms') == 'true'
        try:
            password = generate_password(length, use_upper, use_nums, use_syms)
            return JsonResponse({'result': password})
        except Exception as e:
            return JsonResponse({'error': str(e)}, status=500)


    # ── Name Generator ──
    if tool_slug == 'name-generator':
        count = request.POST.get('count', 10)
        gender = request.POST.get('gender', 'both')
        category = request.POST.get('category', 'person')
        try:
            names = generate_names(count, gender, category)
            return JsonResponse({'result': names})
        except Exception as e:
            return JsonResponse({'error': str(e)}, status=500)

    # ── QR Code Generator ──
    if tool_slug == 'qrcode-generator':
        text = request.POST.get('text', '')
        if not text:
            return JsonResponse({'error': 'Please enter text or a URL.'}, status=400)
        
        # New advanced options (Monkey Features)
        fg_color = request.POST.get('fg_color', '#000000')
        bg_color = request.POST.get('bg_color', '#ffffff')
        style = request.POST.get('style', 'square')
        eye_style = request.POST.get('eye_style', 'square')
        ball_style = request.POST.get('ball_style', 'square')
        gradient = request.POST.get('gradient', 'none')
        output_format = request.POST.get('output_format', 'png')
        
        logo_path = None
        # Check for uploaded logo OR preset logo
        if 'logo' in request.FILES:
            logo_path = save_uploaded_file(request.FILES['logo'])

        try:
            output_path = generate_qr_code(
                text, fg_color=fg_color, bg_color=bg_color, 
                style=style, gradient_type=gradient, 
                eye_style=eye_style, ball_style=ball_style,
                logo_path=logo_path, output_format=output_format
            )
            # Cleanup logo if used
            if logo_path and os.path.exists(logo_path):
                try: os.remove(logo_path)
                except: pass
            
            ct = 'image/jpeg' if output_format.lower() in ('jpg', 'jpeg') else 'image/png'
            return create_cleanup_response(output_path, content_type=ct)
        except Exception as e:
            return JsonResponse({'error': f"QR Generation Failed: {str(e)}"}, status=500)

    # ── Meme Generator ──
    if tool_slug == 'meme-generator':
        if 'file' not in request.FILES:
            return JsonResponse({'error': 'No image uploaded.'}, status=400)
        top_text = request.POST.get('top_text', '')
        bottom_text = request.POST.get('bottom_text', '')
        try:
            uploaded_file = request.FILES['file']
            input_path = save_uploaded_file(uploaded_file)
            output_path = generate_meme(input_path, uploaded_file.name, top_text, bottom_text)
            os.remove(input_path)
            return create_cleanup_response(output_path, content_type='image/jpeg')
        except Exception as e:
            return JsonResponse({'error': str(e)}, status=500)

    # ── Video Downloader (YT/IG) ──
    if tool_slug in ['youtube-downloader', 'instagram-downloader']:
        action = request.POST.get('action', 'info')
        url = request.POST.get('url', '')
        if not url:
            return JsonResponse({'error': 'Please enter a URL.'}, status=400)
        
        try:
            if action == 'info':
                info = get_video_info(url)
                return JsonResponse({'result': info})
            elif action == 'download':
                format_id = request.POST.get('format_id')
                output_path = download_video(url, format_id)
                return create_cleanup_response(output_path, content_type='video/mp4')
        except Exception as e:
            return JsonResponse({'error': str(e)}, status=500)

    # ── Speed Test ──
    if tool_slug == 'speed-test':
        try:
            results = run_speed_test()
            return JsonResponse({'success': True, 'results': results})
        except Exception as e:
            return JsonResponse({'success': False, 'error': str(e)}, status=500)

    # ── AI: Image/Text to Video ──
    if tool_slug == 'image-to-video':
        prompt = request.POST.get('prompt')
        if not prompt:
            return JsonResponse({'error': 'Please provide a prompt for video generation.'}, status=400)
        
        uploaded_file = request.FILES.get('file')
        input_path = None
        if uploaded_file:
            input_path = save_uploaded_file(uploaded_file)
        
        try:
            output_path = generate_video(prompt, input_path=input_path, original_name=uploaded_file.name if uploaded_file else "gen_video")
            
            if input_path and os.path.exists(input_path):
                os.remove(input_path)
                
            return create_cleanup_response(output_path, content_type='video/mp4')
        except Exception as e:
            return JsonResponse({'error': str(e)}, status=500)

    # ── AI: Story Generator ──
    if tool_slug == 'story-generator':
        action = request.POST.get('action', 'info')
        if action == 'info':
            genre = request.POST.get('genre', 'Science Fiction')
            prompt = request.POST.get('prompt', '')
            try:
                story = generate_story(genre, prompt=prompt)
                return JsonResponse({'result': story})
            except Exception as e:
                return JsonResponse({'error': str(e)}, status=500)
        elif action == 'download':
            story_html = request.POST.get('story', '')
            try:
                # Wrap story in professional PDF template
                styled_html = f"""
                <html>
                <head>
                    <style>
                        @page {{ size: A5; margin: 2cm; }}
                        body {{ font-family: serif; line-height: 1.6; color: #333; }}
                        h1 {{ text-align: center; color: #4f46e5; border-bottom: 2px solid #4f46e5; }}
                        .footer {{ text-align: center; font-size: 8pt; color: #999; margin-top: 2cm; }}
                    </style>
                </head>
                <body>
                    <h1>A ScanPDF Story</h1>
                    <div>{story_html}</div>
                    <div class="footer">Generated by ScanPDF AI Story Engine • {time.strftime('%Y')}</div>
                </body>
                </html>
                """
                import weasyprint
                output_path = get_output_path("AI_Story", "pdf")
                weasyprint.HTML(string=styled_html).write_pdf(output_path)
                return create_cleanup_response(output_path, content_type='application/pdf')
            except Exception as e:
                return JsonResponse({'error': f"Failed to generate PDF: {str(e)}"}, status=500)


    # ── Sign PDF ──
    if tool_slug == 'sign-pdf':
        if 'file' not in request.FILES:
            return JsonResponse({'error': 'No file was uploaded.'}, status=400)

        uploaded_file = request.FILES['file']
        signature_data = request.POST.get('signature_data', '')
        page_number = request.POST.get('page_number', '0')
        sig_x = request.POST.get('sig_x', '100')
        sig_y = request.POST.get('sig_y', '600')
        sig_width = request.POST.get('sig_width', '200')
        sig_height = request.POST.get('sig_height', '80')

        signature_image = request.FILES.get('signature_image')
        sig_image_path = None
        if signature_image:
            sig_image_path = save_uploaded_file(signature_image)

        if not signature_data and not sig_image_path:
            return JsonResponse({'error': 'Please provide a signature (draw or upload).'}, status=400)

        try:
            input_path = save_uploaded_file(uploaded_file)
            output_path = sign_pdf(
                input_path, uploaded_file.name,
                signature_image_path=sig_image_path,
                signature_data=signature_data,
                page_number=page_number,
                x=sig_x, y=sig_y,
                width=sig_width, height=sig_height,
            )
            try:
                os.remove(input_path)
            except OSError:
                pass
            if sig_image_path:
                try:
                    os.remove(sig_image_path)
                except OSError:
                    pass
            return create_cleanup_response(output_path, content_type='application/pdf')
        except Exception as e:
            return JsonResponse({'error': f'Sign PDF failed: {str(e)}'}, status=500)

    # ── Redact PDF ──
    if tool_slug == 'redact-pdf':
        if 'file' not in request.FILES:
            return JsonResponse({'error': 'No file was uploaded.'}, status=400)

        uploaded_file = request.FILES['file']
        redaction_areas = request.POST.get('redaction_areas', '[]')

        try:
            input_path = save_uploaded_file(uploaded_file)
            output_path = redact_pdf(
                input_path, uploaded_file.name,
                redaction_areas=redaction_areas,
            )
            try:
                os.remove(input_path)
            except OSError:
                pass
            return create_cleanup_response(output_path, content_type='application/pdf')
        except Exception as e:
            return JsonResponse({'error': f'Redact PDF failed: {str(e)}'}, status=500)

    # ── Default Fallback for other tools ──
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

        # Generate a proper filename based on the tool
        base = Path(uploaded_file.name).stem
        ext = Path(output_path).suffix
        slug_to_suffix = {
            'word-to-pdf': '_converted.pdf',
            'pptx-to-pdf': '_converted.pdf',
            'excel-to-pdf': '_converted.pdf',
            'pdf-to-image': ext,
            'pdf-to-word': '_converted.docx',
            'pdf-to-pptx': '_converted.pptx',
            'pdf-to-excel': '_converted.xlsx',
            'compress-pdf': '_compressed.pdf',
            'pdf-to-pdfa': '_pdfa.pdf',
        }
        suffix = slug_to_suffix.get(tool_slug, f'_output{ext}')
        download_name = f"{base}{suffix}"

        return create_cleanup_response(output_path, content_type=content_type, filename=download_name)

    except Exception as e:
        return JsonResponse({
            'error': f'Conversion failed: {str(e)}'
        }, status=500)

# ─── Speed Test Endpoints ─────────────────────────────────────
@csrf_exempt
def get_client_info(request):
    """Retrieve client and server information for the speed test."""
    x_forwarded_for = request.META.get('HTTP_X_FORWARDED_FOR')
    ip = x_forwarded_for.split(',')[0] if x_forwarded_for else request.META.get('REMOTE_ADDR')
    
    # Generic fallback if no external API reach
    client_data = {
        'ip': ip,
        'city': 'Detected',
        'country_code': 'Network',
        'org': 'Standard ISP'
    }
    
    # Attempt to get real metadata from ipapi securely on server-side (no CORS)
    try:
        import requests
        resp = requests.get(f"https://ipapi.co/{ip}/json/", timeout=3)
        if resp.status_code == 200:
            client_data = resp.json()
    except:
        pass
        
    return JsonResponse(client_data)

@csrf_exempt
def speedtest_download(request):
    """Fast endpoint for testing download speed."""
    from django.http import HttpResponse
    # Increased chunk to 10MB to saturate high-speed connections better
    data = b'0' * (1024 * 1024 * 10) 
    response = HttpResponse(data, content_type='application/octet-stream')
    response['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
    response['Access-Control-Allow-Origin'] = '*'
    return response

@csrf_exempt
def speedtest_upload(request):
    """Endpoint for testing upload speed."""
    if request.method == 'POST':
        # Consume the data to measure upload time
        _ = request.body
    return JsonResponse({'success': True})


def custom_404_view(request, exception=None):
    """Custom view for handling 404 errors."""
    return render(request, '404.html', status=404)
