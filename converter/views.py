"""
Views for the file converter application.
"""
import os
import mimetypes
import json
import time
import urllib.request
import urllib.parse
from pathlib import Path
from django.shortcuts import render
from django.http import JsonResponse, FileResponse, Http404, HttpResponse
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
    balance_chemical_equation,
    generate_qr_code,
    generate_meme,
    generate_password,
    generate_story,
    generate_names,
    run_speed_test,
    convert_images_to_pdf,
    convert_pdf_to_pdfa,
    sign_pdf,
    redact_pdf,
    merge_word_files,
)
from .utils_video import convert_video_format

_CURRENCY_CACHE = {'base': None, 'rates': None, 'updated_at': 0}


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

    from .utils import format_download_name

    final_filename = format_download_name(filename or os.path.basename(file_path), file_path)

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
        'color': '#7c3aed',
        'gradient': 'from-violet-600 to-purple-900',
        'category': 'convert',

        'seo_title': 'Word to PDF Converter Online – Free DOCX to PDF',
        'seo_description': 'Convert Word documents to PDF online for free. Upload a DOCX file and quickly create a professional PDF without registration.',
        'seo_keywords': 'word to pdf, docx to pdf, convert word to pdf, word pdf converter, free word to pdf',
        'seo_h1': 'Word to PDF Converter',
        's1': 'Word to',
        'highlight': 'PDF Converter',
        'seo_intro': 'Convert Word documents to PDF online quickly and easily. Upload your DOCX file and create a professional PDF without installing software or creating an account.',
    },

    'pptx-to-pdf': {
        'title': 'PowerPoint to PDF',
        'description': 'Convert PowerPoint presentations to PDF files quickly and easily.',
        'icon': 'presentation',
        'accept': '.pptx',
        'allowed_extensions': ['.pptx'],
        'converter': convert_pptx_to_pdf,
        'color': '#c05621',
        'gradient': 'from-orange-500 to-red-500',
        'category': 'convert',

        'seo_title': 'PowerPoint to PDF Converter Online – Free PPTX to PDF',
        'seo_description': 'Convert PowerPoint PPTX presentations to PDF online for free. Preserve your presentation content and create an easy-to-share PDF document.',
        'seo_keywords': 'powerpoint to pdf, pptx to pdf, convert pptx to pdf, powerpoint pdf converter, free ppt to pdf',
        'seo_h1': 'PowerPoint to PDF Converter',
        's1': 'PowerPoint to',
        'highlight': 'PDF Converter',
        'seo_intro': 'Convert PowerPoint presentations to PDF online. Upload your PPTX file and quickly create a professional PDF that is easy to share and view.',
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

        'seo_title': 'Excel to PDF Converter Online – Free XLSX to PDF',
        'seo_description': 'Convert Excel XLSX spreadsheets to PDF online for free. Turn your spreadsheets into clean, professional PDF documents in seconds.',
        'seo_keywords': 'excel to pdf, xlsx to pdf, convert excel to pdf, excel pdf converter, free excel to pdf',
        'seo_h1': 'Excel to PDF Converter',
        's1': 'Excel to',
        'highlight': 'PDF Converter',
        'seo_intro': 'Convert Excel spreadsheets to PDF online without complicated software. Upload your XLSX file and create a clean, shareable PDF document.',
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

        'seo_title': 'HTML to PDF Converter Online – Free HTML to PDF',
        'seo_description': 'Convert HTML files and web pages to PDF online. Create high-quality PDF documents from HTML quickly and easily.',
        'seo_keywords': 'html to pdf, html converter, convert html to pdf, webpage to pdf, html pdf converter',
        'seo_h1': 'HTML to PDF Converter',
        's1': 'HTML to',
        'highlight': 'PDF Converter',
        'seo_intro': 'Convert HTML files into professional PDF documents online. Create a PDF from your HTML content while maintaining a clean layout and formatting.',
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

        'seo_title': 'PDF to JPG & PNG Converter Online – Free PDF to Image',
        'seo_description': 'Convert PDF pages to high-quality JPG or PNG images online for free. Extract individual PDF pages as image files quickly.',
        'seo_keywords': 'pdf to image, pdf to jpg, pdf to png, convert pdf to image, pdf image converter',
        'seo_h1': 'PDF to Image Converter',
        's1': 'PDF to',
        'highlight': 'Image Converter',
        'seo_intro': 'Convert PDF pages into JPG or PNG images online. Upload your PDF, choose your preferred image format, and download high-quality images.',
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

        'seo_title': 'PDF to Word Converter Online – Free PDF to DOCX',
        'seo_description': 'Convert PDF files to editable Word DOCX documents online for free. Quickly extract PDF content into an editable Word file.',
        'seo_keywords': 'pdf to word, pdf to docx, convert pdf to word, pdf word converter, free pdf to word',
        'seo_h1': 'PDF to Word Converter',
        's1': 'PDF to',
        'highlight': 'Word Converter',
        'seo_intro': 'Convert PDF documents into editable Word files online. Upload your PDF and create a DOCX document that you can edit and reuse.',
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

        'seo_title': 'PDF to PowerPoint Converter Online – Free PDF to PPTX',
        'seo_description': 'Convert PDF files to editable PowerPoint PPTX presentations online. Quickly transform PDF content into a presentation format.',
        'seo_keywords': 'pdf to powerpoint, pdf to pptx, convert pdf to ppt, pdf ppt converter, free pdf to powerpoint',
        'seo_h1': 'PDF to PowerPoint Converter',
        's1': 'PDF to',
        'highlight': 'PowerPoint Converter',
        'seo_intro': 'Convert PDF documents into PowerPoint presentations online. Upload your PDF and transform it into an editable PPTX file.',
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

        'seo_title': 'PDF to Excel Converter Online – Free PDF to XLSX',
        'seo_description': 'Convert PDF tables into editable Excel XLSX spreadsheets online. Extract structured data from PDF files quickly and easily.',
        'seo_keywords': 'pdf to excel, pdf to xlsx, convert pdf to excel, pdf excel converter, extract pdf table',
        'seo_h1': 'PDF to Excel Converter',
        's1': 'PDF to',
        'highlight': 'Excel Converter',
        'seo_intro': 'Convert PDF tables and data into editable Excel spreadsheets. Upload your PDF and create an XLSX workbook from your document.',
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

        'seo_title': 'Merge PDF Files Online – Free PDF Merger',
        'seo_description': 'Merge multiple PDF files into one document online for free. Combine PDFs in your preferred order without registration.',
        'seo_keywords': 'merge pdf, combine pdf, pdf merger, merge pdf files, combine pdf online',
        'seo_h1': 'Merge PDF Files Online',
        's1': 'Merge',
        'highlight': 'PDF Files Online',
        'seo_intro': 'Combine multiple PDF files into one organized document. Upload your PDFs, arrange them in the desired order, and merge them into a single file.',
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

        'seo_title': 'Split PDF Online – Free PDF Splitter',
        'seo_description': 'Split PDF files into individual pages or custom page ranges online for free. Extract the pages you need quickly.',
        'seo_keywords': 'split pdf, pdf splitter, split pdf pages, extract pdf pages, split pdf online',
        'seo_h1': 'Split PDF Online',
        's1': 'Split',
        'highlight': 'PDF Online',
        'seo_intro': 'Split PDF documents into individual pages or custom page ranges. Upload your PDF and extract exactly the pages you need.',
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

        'seo_title': 'Compress PDF Online – Free PDF Compressor',
        'seo_description': 'Reduce PDF file size online while maintaining good visual quality. Compress large PDF documents quickly and easily.',
        'seo_keywords': 'compress pdf, pdf compressor, reduce pdf size, compress pdf online, shrink pdf',
        'seo_h1': 'Compress PDF Online',
        's1': 'Compress',
        'highlight': 'PDF Online',
        'seo_intro': 'Reduce the size of large PDF documents without unnecessary complexity. Upload your PDF and create a smaller, easier-to-share file.',
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

        'seo_title': 'Remove Pages from PDF Online – Free PDF Page Remover',
        'seo_description': 'Remove unwanted pages from PDF files online. Delete specific PDF pages quickly and create a clean document.',
        'seo_keywords': 'remove pages from pdf, delete pdf pages, pdf page remover, remove pdf page',
        'seo_h1': 'Remove Pages from PDF',
        's1': 'Remove Pages from',
        'highlight': 'PDF',
        'seo_intro': 'Delete unwanted pages from your PDF document. Select the pages you want to remove and create a cleaner PDF file.',
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

        'seo_title': 'Extract Pages from PDF Online – Free PDF Page Extractor',
        'seo_description': 'Extract selected pages from PDF files and save them as a new PDF document online for free.',
        'seo_keywords': 'extract pages from pdf, pdf page extractor, extract pdf pages, save pdf pages',
        'seo_h1': 'Extract Pages from PDF',
        's1': 'Extract Pages from',
        'highlight': 'PDF',
        'seo_intro': 'Extract specific pages from a PDF and create a new document containing only the pages you need.',
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

        'seo_title': 'Organize PDF Pages Online – Free PDF Page Organizer',
        'seo_description': 'Reorder and organize PDF pages online. Rearrange your document pages into the correct order quickly and easily.',
        'seo_keywords': 'organize pdf, reorder pdf pages, rearrange pdf pages, pdf page organizer',
        'seo_h1': 'Organize PDF Pages',
        's1': 'Organize',
        'highlight': 'PDF Pages',
        'seo_intro': 'Rearrange PDF pages into the order you want. Organize your document quickly and create a properly structured PDF.',
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

        'seo_title': 'Repair PDF Online – Free PDF Repair Tool',
        'seo_description': 'Try to repair corrupted or damaged PDF files online and recover accessible document content.',
        'seo_keywords': 'repair pdf, fix corrupted pdf, damaged pdf repair, pdf repair tool, recover pdf',
        'seo_h1': 'Repair PDF Online',
        's1': 'Repair',
        'highlight': 'PDF Online',
        'seo_intro': 'Attempt to repair damaged or corrupted PDF documents and recover usable content from broken files.',
    },

    'ocr-pdf': {
        'title': 'OCR to PDF',
        'description': 'Convert scanned PDFs, Word documents, and images into searchable, selectable PDF documents or extract their text directly.',
        'icon': 'languages',
        'accept': '.pdf,.jpg,.jpeg,.png,.docx',
        'allowed_extensions': ['.pdf', '.jpg', '.jpeg', '.png', '.docx'],
        'converter': ocr_pdf,
        'color': '#0d9488',
        'gradient': 'from-teal-500 to-teal-700',
        'category': 'pdf-tools',
        'multi_file': True,

        'seo_title': 'OCR PDF Online – Convert Scanned Documents to Searchable PDF',
        'seo_description': 'Use OCR to convert scanned PDFs, images, and documents into searchable and selectable PDF files or extract text online.',
        'seo_keywords': 'ocr pdf, pdf ocr, scanned pdf to text, image to text, searchable pdf, ocr online',
        'seo_h1': 'OCR PDF Converter',
        's1': 'OCR',
        'highlight': 'PDF Converter',
        'seo_intro': 'Convert scanned documents and images into searchable PDF files using OCR technology. Extract text from scanned PDFs, JPGs, PNGs, and supported documents.',
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

        'seo_title': 'Rotate PDF Online – Free PDF Page Rotator',
        'seo_description': 'Rotate PDF pages online by 90, 180, or 270 degrees. Fix incorrectly oriented PDF pages quickly.',
        'seo_keywords': 'rotate pdf, rotate pdf pages, pdf rotator, rotate pdf online',
        'seo_h1': 'Rotate PDF Online',
        's1': 'Rotate',
        'highlight': 'PDF Online',
        'seo_intro': 'Rotate PDF pages to the correct orientation. Quickly turn individual pages or entire documents by 90, 180, or 270 degrees.',
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

        'seo_title': 'Add Watermark to PDF Online – Free PDF Watermark Tool',
        'seo_description': 'Add custom text watermarks to PDF documents online. Protect and brand your PDF files with a personalized watermark.',
        'seo_keywords': 'add watermark to pdf, pdf watermark, watermark pdf online, pdf watermark tool',
        'seo_h1': 'Add Watermark to PDF',
        's1': 'Add Watermark to',
        'highlight': 'PDF',
        'seo_intro': 'Add a custom text watermark to your PDF documents. Use watermarks for branding, identification, or document labeling.',
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

        'seo_title': 'Remove Watermark from PDF Online – PDF Watermark Remover',
        'seo_description': 'Attempt to remove watermarks from PDF documents online. Upload a PDF and process watermark removal.',
        'seo_keywords': 'remove watermark pdf, pdf watermark remover, erase pdf watermark, remove watermark online',
        'seo_h1': 'Remove Watermark from PDF',
        's1': 'Remove Watermark from',
        'highlight': 'PDF',
        'seo_intro': 'Attempt to remove watermarks from PDF documents using an online PDF watermark removal tool.',
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

        'seo_title': 'Crop PDF Online – Free PDF Page Cropper',
        'seo_description': 'Crop PDF pages and remove unwanted whitespace or margins online. Resize PDF page areas quickly and easily.',
        'seo_keywords': 'crop pdf, crop pdf pages, pdf cropper, trim pdf, crop pdf online',
        'seo_h1': 'Crop PDF Online',
        's1': 'Crop',
        'highlight': 'PDF Online',
        'seo_intro': 'Crop PDF pages to remove unwanted whitespace and adjust page margins for a cleaner document layout.',
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

        'seo_title': 'Edit PDF Online – Free PDF Editor',
        'seo_description': 'Edit PDF documents online by adding text, annotations, and notes to your PDF pages.',
        'seo_keywords': 'edit pdf, pdf editor, edit pdf online, annotate pdf, pdf annotation tool',
        'seo_h1': 'Edit PDF Online',
        's1': 'Edit',
        'highlight': 'PDF Online',
        'seo_intro': 'Make quick changes to PDF documents by adding text annotations and notes directly to your PDF pages.',
    },

    'unlock-pdf': {
        'title': 'Unlock PDF',
        'description': 'Remove password protection from your secured PDF files.',
        'icon': 'unlock',
        'accept': '.pdf',
        'allowed_extensions': ['.pdf'],
        'converter': None,
        'color': '#10b981',
        'gradient': 'from-emerald-500 to-emerald-700',
        'category': 'pdf-tools',

        'seo_title': 'Unlock PDF Online – Free PDF Password Remover',
        'seo_description': 'Unlock password-protected PDF files online when you have authorization to access the document.',
        'seo_keywords': 'unlock pdf, remove pdf password, pdf password remover, unlock pdf online',
        'seo_h1': 'Unlock PDF Online',
        's1': 'Unlock',
        'highlight': 'PDF Online',
        'seo_intro': 'Unlock secured PDF documents when you have permission to access them. Remove supported PDF restrictions and make your document accessible.',
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

        'seo_title': 'Protect PDF with Password Online – Free PDF Security Tool',
        'seo_description': 'Protect PDF documents with password encryption online. Add security to your PDF files and restrict unauthorized access.',
        'seo_keywords': 'protect pdf, password protect pdf, encrypt pdf, pdf password, secure pdf',
        'seo_h1': 'Protect PDF with Password',
        's1': 'Protect PDF with',
        'highlight': 'Password',
        'seo_intro': 'Add password protection to PDF documents to help restrict unauthorized access and keep sensitive files secure.',
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

        'seo_title': 'PNG to JPG Converter Online – Free PNG to JPEG',
        'seo_description': 'Convert PNG images to JPG format online for free. Reduce image file size and create compatible JPEG images quickly.',
        'seo_keywords': 'png to jpg, png to jpeg, convert png to jpg, png jpg converter',
        'seo_h1': 'PNG to JPG Converter',
        's1': 'PNG to',
        'highlight': 'JPG Converter',
        'seo_intro': 'Convert PNG images to JPG online. Upload your PNG file and quickly create a JPEG image suitable for websites, sharing, and storage.',
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

        'seo_title': 'JPG to PNG Converter Online – Free JPEG to PNG',
        'seo_description': 'Convert JPG and JPEG images to PNG format online for free. Create PNG images for editing, transparency, and high-quality graphics.',
        'seo_keywords': 'jpg to png, jpeg to png, convert jpg to png, jpg png converter',
        'seo_h1': 'JPG to PNG Converter',
        's1': 'JPG to',
        'highlight': 'PNG Converter',
        'seo_intro': 'Convert JPG and JPEG images to PNG online. Create high-quality PNG files for editing, graphics, and images that require transparency.',
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

        'seo_title': 'HTML to Image Converter Online – Free HTML Screenshot Tool',
        'seo_description': 'Convert HTML files into high-quality images online. Capture HTML content as an image while preserving its visual layout.',
        'seo_keywords': 'html to image, html screenshot, convert html to image, html image converter',
        'seo_h1': 'HTML to Image Converter',
        's1': 'HTML to',
        'highlight': 'Image Converter',
        'seo_intro': 'Convert HTML content into an image online. Create a visual snapshot of your HTML file while preserving its page layout.',
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

        'seo_title': 'Resize Image Online – Free JPG PNG Image Resizer',
        'seo_description': 'Resize JPG, JPEG, and PNG images online. Set custom image dimensions quickly while preparing images for websites or sharing.',
        'seo_keywords': 'resize image, image resizer, resize jpg, resize png, image resize online',
        'seo_h1': 'Resize Image Online',
        's1': 'Resize',
        'highlight': 'Image Online',
        'seo_intro': 'Resize JPG, JPEG, and PNG images to custom dimensions. Enter your preferred width and height and create a properly sized image.',
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

        'seo_title': 'Scale Image Online – Free Image Scaling Tool',
        'seo_description': 'Scale JPG and PNG images up or down by percentage online. Quickly change image dimensions while maintaining proportions.',
        'seo_keywords': 'scale image, image scaling, enlarge image, reduce image size, scale jpg png',
        'seo_h1': 'Scale Image Online',
        's1': 'Scale',
        'highlight': 'Image Online',
        'seo_intro': 'Scale images up or down using a percentage-based resize tool. Quickly adjust image dimensions while maintaining the desired proportions.',
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

        'seo_title': 'Rotate Image Online – Free JPG PNG Image Rotator',
        'seo_description': 'Rotate JPG, JPEG, and PNG images online by any angle. Correct image orientation quickly with a simple image rotator.',
        'seo_keywords': 'rotate image, image rotator, rotate jpg, rotate png, rotate image online',
        'seo_h1': 'Rotate Image Online',
        's1': 'Rotate',
        'highlight': 'Image Online',
        'seo_intro': 'Rotate JPG, JPEG, and PNG images by your preferred angle. Correct image orientation quickly without installing image editing software.',
    },

    'watermark-image': {
        'title': 'Add Watermark',
        'description': 'Overlay a custom text watermark on your images.',
        'icon': 'stamp',
        'accept': '.jpg,.jpeg,.png',
        'allowed_extensions': ['.jpg', '.jpeg', '.png'],
        'converter': None,
        'color': '#0ea5e9',
        'gradient': 'from-sky-500 to-sky-700',
        'category': 'image-tools',

        'seo_title': 'Add Watermark to Image Online – Free Image Watermark Tool',
        'seo_description': 'Add custom text watermarks to JPG and PNG images online. Protect, label, or brand your images quickly.',
        'seo_keywords': 'watermark image, add watermark to image, image watermark, watermark jpg png',
        'seo_h1': 'Add Watermark to Image',
        's1': 'Add Watermark to',
        'highlight': 'Image',
        'seo_intro': 'Add a custom text watermark to your images for branding, identification, or protection. Upload your image and apply a personalized watermark.',
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

        'seo_title': 'Compress Image Online – Free JPG PNG Image Compressor',
        'seo_description': 'Compress JPG, JPEG, and PNG images online to reduce file size while maintaining good image quality.',
        'seo_keywords': 'compress image, image compressor, compress jpg, compress png, reduce image size',
        'seo_h1': 'Compress Image Online',
        's1': 'Compress',
        'highlight': 'Image Online',
        'seo_intro': 'Reduce JPG and PNG image file sizes while maintaining good visual quality. Compress images for websites, email, and faster sharing.',
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

        'seo_title': 'Crop Image Online – Free JPG PNG Image Cropper',
        'seo_description': 'Crop JPG and PNG images online with a precise rectangular selection. Remove unwanted areas and create the perfect image size.',
        'seo_keywords': 'crop image, image cropper, crop jpg, crop png, crop image online',
        'seo_h1': 'Crop Image Online',
        's1': 'Crop',
        'highlight': 'Image Online',
        'seo_intro': 'Crop JPG, JPEG, and PNG images to remove unwanted areas or create the exact composition you need.',
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

        'seo_title': 'Chemical Equation Balancer Online – Free Chemistry Tool',
        'seo_description': 'Balance chemical equations online using stoichiometry. Quickly find balanced coefficients for common chemical reactions.',
        'seo_keywords': 'chemical equation balancer, balance chemical equations, chemistry calculator, stoichiometry calculator',
        'seo_h1': 'Chemical Equation Balancer',
        's1': 'Chemical Equation',
        'highlight': 'Balancer',
        'seo_intro': 'Balance chemical equations online using stoichiometry. Enter a chemical reaction and calculate the correct coefficients.',
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

        'seo_title': 'Strong Password Generator Online – Free Secure Password Tool',
        'seo_description': 'Generate strong random passwords online for free. Create secure passwords using customizable length and character options.',
        'seo_keywords': 'password generator, strong password generator, secure password generator, random password generator',
        'seo_h1': 'Strong Password Generator',
        's1': 'Strong Password',
        'highlight': 'Generator',
        'seo_intro': 'Create strong, random passwords online. Customize password length and character types to generate secure passwords for your accounts.',
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

        'seo_title': 'Unit Converter Online – Length, Weight, Temperature & More',
        'seo_description': 'Convert units online including length, weight, temperature, volume, and more with a simple free unit converter.',
        'seo_keywords': 'unit converter, online unit converter, length converter, weight converter, temperature converter',
        'seo_h1': 'Online Unit Converter',
        's1': 'Online Unit',
        'highlight': 'Converter',
        'seo_intro': 'Convert common units for length, weight, temperature, volume, and more using a fast and easy online unit converter.',
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

        'seo_title': 'Internet Speed Test Online – Check Download & Upload Speed',
        'seo_description': 'Test your internet connection speed online. Check download speed, upload speed, latency, and connection information.',
        'seo_keywords': 'internet speed test, wifi speed test, broadband speed test, download speed test, upload speed test',
        'seo_h1': 'Internet Speed Test',
        's1': 'Internet',
        'highlight': 'Speed Test',
        'seo_intro': 'Check your internet connection performance with an online speed test. Measure download speed, upload speed, and connection details.',
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

        'seo_title': 'QR Code Generator Online – Free Custom QR Code Maker',
        'seo_description': 'Create QR codes online for URLs, text, Wi-Fi, and more. Generate custom QR codes quickly without registration.',
        'seo_keywords': 'qr code generator, free qr code generator, qr code maker, create qr code, online qr generator',
        'seo_h1': 'Free QR Code Generator',
        's1': 'Free QR Code',
        'highlight': 'Generator',
        'seo_intro': 'Create QR codes online for websites, text, Wi-Fi information, and more. Generate a QR code quickly and easily.',
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

        'seo_title': 'Meme Generator Online – Create Free Custom Memes',
        'seo_description': 'Create custom memes online by adding text to your JPG and PNG images. Make and download memes quickly.',
        'seo_keywords': 'meme generator, meme maker, create meme online, free meme generator, custom meme maker',
        'seo_h1': 'Free Meme Generator',
        's1': 'Free Meme',
        'highlight': 'Generator',
        'seo_intro': 'Create custom memes online by uploading an image and adding your own text. Make funny memes quickly and easily.',
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

        'seo_title': 'Random Name Generator Online – Free Name Ideas',
        'seo_description': 'Generate random name ideas for people, places, businesses, characters, and creative projects with a free online name generator.',
        'seo_keywords': 'name generator, random name generator, business name generator, character name generator',
        'seo_h1': 'Random Name Generator',
        's1': 'Random Name',
        'highlight': 'Generator',
        'seo_intro': 'Generate creative name ideas for people, places, companies, characters, and other projects using a simple online name generator.',
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

        'seo_title': 'AI Story Generator Online – Create Stories with AI',
        'seo_description': 'Generate creative stories online with AI. Create original fiction, characters, plots, and stories across different genres.',
        'seo_keywords': 'ai story generator, story generator, ai writing tool, generate stories with ai, free story generator',
        'seo_h1': 'AI Story Generator',
        's1': 'AI Story',
        'highlight': 'Generator',
        'seo_intro': 'Create original stories with AI. Choose a genre or idea and generate creative plots, characters, and storytelling content online.',
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

        'seo_title': 'Image to PDF Converter Online – JPG PNG to PDF',
        'seo_description': 'Convert JPG, JPEG, and PNG images to PDF online for free. Combine multiple images into a single PDF document.',
        'seo_keywords': 'image to pdf, jpg to pdf, png to pdf, convert image to pdf, photo to pdf',
        'seo_h1': 'Image to PDF Converter',
        's1': 'Image to',
        'highlight': 'PDF Converter',
        'seo_intro': 'Convert JPG, JPEG, and PNG images into a single PDF document. Upload one or multiple images and create a PDF quickly.',
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

        'seo_title': 'PDF to PDF/A Converter Online – Archival PDF Conversion',
        'seo_description': 'Convert PDF documents to PDF/A archival format for long-term preservation, document archiving, and compliance workflows.',
        'seo_keywords': 'pdf to pdfa, pdf to pdf/a, pdfa converter, archival pdf, pdf archival format',
        'seo_h1': 'PDF to PDF/A Converter',
        's1': 'PDF to',
        'highlight': 'PDF/A Converter',
        'seo_intro': 'Convert PDF documents to PDF/A archival format for long-term document preservation and compatible archival workflows.',
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

        'seo_title': 'Sign PDF Online – Free PDF Signature Tool',
        'seo_description': 'Sign PDF documents online by drawing, typing, or adding a signature image to your PDF pages.',
        'seo_keywords': 'sign pdf, pdf signature, electronic signature pdf, sign pdf online, add signature to pdf',
        'seo_h1': 'Sign PDF Online',
        's1': 'Sign',
        'highlight': 'PDF Online',
        'seo_intro': 'Add a signature to your PDF document online. Draw, type, or upload your signature and place it on the required PDF page.',
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

        'seo_title': 'Redact PDF Online – Free PDF Redaction Tool',
        'seo_description': 'Redact sensitive information from PDF documents by permanently covering confidential text and areas.',
        'seo_keywords': 'redact pdf, pdf redaction, redact pdf online, hide sensitive information pdf',
        'seo_h1': 'Redact PDF Online',
        's1': 'Redact',
        'highlight': 'PDF Online',
        'seo_intro': 'Remove sensitive information from PDF documents by redacting confidential text and areas before sharing the document.',
    },

    'audio-editor': {
        'title': 'Audio Editor',
        'description': 'A professional-grade audio editing suite to trim, change volume, speed, pitch and apply equalizer effects.',
        'icon': 'music',
        'accept': '.mp3,.wav,.ogg,.m4a,.flac',
        'allowed_extensions': ['.mp3', '.wav', '.ogg', '.m4a', '.flac'],
        'converter': None,
        'color': '#10b981',
        'gradient': 'from-emerald-500 to-teal-600',
        'category': 'audio-tools',

        'seo_title': 'Online Audio Editor – Free MP3 WAV Audio Editing Tool',
        'seo_description': 'Edit audio online with tools for trimming, volume, speed, pitch, and equalizer effects. Supports common audio formats.',
        'seo_keywords': 'audio editor, online audio editor, mp3 editor, wav editor, trim audio, audio editing tool',
        'seo_h1': 'Online Audio Editor',
        's1': 'Online',
        'highlight': 'Audio Editor',
        'seo_intro': 'Edit audio files online with tools for trimming, volume adjustment, speed, pitch, and equalizer effects.',
    },

    'merge-audio': {
        'title': 'Merge Audio',
        'description': 'Combine multiple audio files into one track with in-page preview and processing status.',
        'icon': 'combine',
        'accept': '.mp3,.wav,.ogg,.m4a,.flac',
        'allowed_extensions': ['.mp3', '.wav', '.ogg', '.m4a', '.flac'],
        'converter': None,
        'color': '#059669',
        'gradient': 'from-emerald-500 to-green-700',
        'category': 'audio-tools',
        'multi_file': True,

        'seo_title': 'Merge Audio Files Online – Free Audio Merger',
        'seo_description': 'Combine multiple MP3, WAV, OGG, M4A, or FLAC audio files into one track online.',
        'seo_keywords': 'merge audio, merge mp3, combine audio files, audio merger, join mp3 files',
        'seo_h1': 'Merge Audio Files Online',
        's1': 'Merge',
        'highlight': 'Audio Files Online',
        'seo_intro': 'Combine multiple audio files into a single track. Upload your audio files, arrange them in order, and merge them into one file.',
    },

    'extract-audio-from-video': {
        'title': 'Extract Audio From Video',
        'description': 'Upload a video, preview it, and extract full audio or a custom start-to-end range.',
        'icon': 'video',
        'accept': '.mp4,.mov,.avi,.mkv,.webm',
        'allowed_extensions': ['.mp4', '.mov', '.avi', '.mkv', '.webm'],
        'converter': None,
        'color': '#2563eb',
        'gradient': 'from-blue-500 to-indigo-700',
        'category': 'audio-tools',

        'seo_title': 'Extract Audio from Video Online – Free Video to Audio Tool',
        'seo_description': 'Extract audio from MP4, MOV, AVI, MKV, and WebM videos online. Choose the full audio or a custom section to extract.',
        'seo_keywords': 'extract audio from video, video to audio, mp4 to mp3, extract mp3 from video, video audio extractor',
        'seo_h1': 'Extract Audio from Video',
        's1': 'Extract Audio from',
        'highlight': 'Video',
        'seo_intro': 'Extract audio tracks from supported video files online. Upload a video and extract the complete audio or a selected time range.',
    },

    'video-converter': {
        'title': 'Video Converter',
        'description': 'Convert videos between formats like MP4, AVI, MOV, MKV, and more.',
        'icon': 'video',
        'accept': '.mp4,.avi,.mov,.mkv,.wmv,.flv,.3gp',
        'allowed_extensions': ['.mp4', '.avi', '.mov', '.mkv', '.wmv', '.flv', '.3gp'],
        'converter': convert_video_format,
        'color': '#7b69a7',
        'gradient': 'from-violet-800 to-purple-900',
        'category': 'convert',

        'seo_title': 'Video Converter Online – Free MP4, AVI, MOV & MKV Converter',
        'seo_description': 'Convert videos online between MP4, AVI, MOV, MKV, WMV, FLV, and 3GP formats with a simple free video converter.',
        'seo_keywords': 'video converter, online video converter, mp4 converter, avi converter, mov converter, mkv converter',
        'seo_h1': 'Online Video Converter',
        's1': 'Online',
        'highlight': 'Video Converter',
        'seo_intro': 'Convert video files between popular formats including MP4, AVI, MOV, MKV, WMV, FLV, and 3GP using an easy online video converter.',
    },
}



def home(request):
    """Render the home page with all available tools."""
    all_tools = {**TOOLS}
    # Update with image tools metadata for the grid
    from image_processor.views import IMAGE_TOOLS
    all_tools.update(IMAGE_TOOLS)
    
    # Auto-discover videos from Django Admin (HeroVideo model)
    from .models import HeroVideo
    
    short_url_videos = []
    qr_code_videos = []
    # Fetch active videos, ordering is already handled by Meta class ("order", "id")
    for video_obj in HeroVideo.objects.filter(is_active=True):
        if video_obj.video:
            if video_obj.section == 'qr_code':
                qr_code_videos.append(video_obj.video.url)
            else:
                short_url_videos.append(video_obj.video.url)
    
    context = {
        'tools': all_tools,
        'page_title': 'ScanPDF',
        'IMAGE_TOOLS_KEYS': list(IMAGE_TOOLS.keys()),
        'hero_videos': short_url_videos,
        'qr_hero_videos': qr_code_videos,
    }
    return render(request, 'converter/home.html', context)


def convert_page(request, tool_slug):
    """Render the conversion page for a specific tool."""
    if tool_slug not in TOOLS:
        raise Http404("Tool not found")

    tool = TOOLS[tool_slug]
    form = FileUploadForm()

    # Determine which template to use
    if tool_slug == 'merge-pdf' or tool_slug == 'merge-word':
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
    elif tool_slug == 'add-image-watermark' or tool_slug == 'watermark-image':
        template = 'converter/add_image_watermark.html'
    elif tool_slug == 'compress-image':
        template = 'converter/compress_image.html'
    elif tool_slug == 'crop-image' or tool_slug == 'cut-image':
        template = 'converter/crop_image.html'

    elif tool_slug == 'chemical-balancer':
        template = 'converter/chemical_balancer.html'
    elif tool_slug == 'password-generator':
        template = 'converter/password_generator.html'
    elif tool_slug == 'unit-converter':
        template = 'converter/unit_converter.html'
    elif tool_slug == 'speed-test':
        template = 'converter/speed_test.html'

    elif tool_slug == 'qrcode-generator':
        template = 'converter/qrcode_generator.html'
    elif tool_slug == 'meme-generator':
        template = 'converter/meme_generator.html'
    elif tool_slug == 'story-generator':
        template = 'converter/story_generator.html'
    elif tool_slug == 'name-generator':
        template = 'converter/name_generator.html'
    elif tool_slug == 'sign-pdf':
        template = 'converter/sign_pdf.html'
    elif tool_slug == 'redact-pdf':
        template = 'converter/redact_pdf.html'
    elif tool_slug == 'html-to-pdf':
        template = 'converter/html_to_pdf.html'
    elif tool_slug == 'audio-editor':
        template = 'audio_processor/editor.html'
    elif tool_slug == 'merge-audio':
        template = 'audio_processor/merge_audio.html'
    elif tool_slug == 'extract-audio-from-video':
        template = 'audio_processor/extract_audio.html'
    elif tool_slug == 'video-converter':
        template = 'converter/video_converter.html'
    else:
        template = 'converter/convert.html'

    context = {
        'tool': tool,
        'tool_slug': tool_slug,
        'form': form,
        'page_title': f'{tool["title"]} — ScanPDF',
                'tool': tool,
        'tool_slug': tool_slug,
        'seo_title': tool.get('seo_title', tool['title']),
        'seo_description': tool.get('seo_description', tool['description']),
        'seo_keywords': tool.get('seo_keywords', ''),
        'seo_h1': tool.get('seo_h1', tool['title']),
        'seo_intro': tool.get('seo_intro', tool['description']),

    's1': tool.get('s1', ''),
    'highlight': tool.get('highlight', ''),
    }
    return render(request, template, context)


@csrf_exempt
@require_POST
def convert_file(request, tool_slug):
    """Handle file conversion via AJAX request."""
    tool = TOOLS[tool_slug]

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
                                           filename=files[0].name)
        except Exception as e:
            return JsonResponse({'error': f'Merge failed: {str(e)}'}, status=500)

    # ── Merge Word: multiple files ──
    if tool_slug == 'merge-word':
        files = request.FILES.getlist('files')
        if not files or len(files) < 2:
            return JsonResponse({'error': 'Please upload at least 2 Word files to merge.'}, status=400)

        try:
            input_paths = []
            for f in files:
                ext = os.path.splitext(f.name)[1].lower()
                if ext != '.docx':
                    return JsonResponse({'error': f'Invalid file "{f.name}". Only .docx files are allowed.'}, status=400)
                input_paths.append(save_uploaded_file(f))

            output_path = merge_word_files(input_paths, files[0].name)

            for p in input_paths:
                try: os.remove(p)
                except OSError: pass

            return create_cleanup_response(output_path, content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
        except Exception as e:
            return JsonResponse({'error': f'Word Merge failed: {str(e)}'}, status=500)

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
                original_name = request.POST.get('original_filename', 'edited_document.pdf')
                return create_cleanup_response(output_path, content_type='application/pdf', filename=original_name)
            else:
                # Case 1: Initial upload - convert PDF to editable HTML
                uploaded_file = request.FILES['file']
                input_path = save_uploaded_file(uploaded_file)
                html_data_list = convert_pdf_to_html_via_word(input_path)
                
                try:
                    os.remove(input_path)
                except OSError:
                    pass
                
                return JsonResponse({
                    'success': True,
                    'pages': html_data_list,
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
    if tool_slug == 'add-image-watermark' or tool_slug == 'watermark-image':
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
    if tool_slug == 'crop-image' or tool_slug == 'cut-image':
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
            
            if output_format.lower() in ('jpg', 'jpeg'):
                ct = 'image/jpeg'
            elif output_format.lower() == 'svg':
                ct = 'image/svg+xml'
            else:
                ct = 'image/png'
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



    # ── Speed Test ──
    if tool_slug == 'speed-test':
        try:
            results = run_speed_test()
            return JsonResponse({'success': True, 'results': results})
        except Exception as e:
            return JsonResponse({'success': False, 'error': str(e)}, status=500)

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
        signatures_json = request.POST.get('signatures_json', '')

        # Legacy single-signature fields (backward compat)
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

        if not signatures_json and not signature_data and not sig_image_path:
            return JsonResponse({'error': 'Please provide a signature (draw or upload).'}, status=400)

        try:
            # If using advanced mode, resolve file-based signature images
            # The frontend sends each sig as 'sig_file_N' in request.FILES
            sig_file_paths = {}
            if signatures_json:
                import json as _json
                placements = _json.loads(signatures_json) if isinstance(signatures_json, str) else signatures_json
                for item in placements:
                    file_key = item.get('file_key', '')
                    if file_key and file_key in request.FILES:
                        saved = save_uploaded_file(request.FILES[file_key])
                        sig_file_paths[file_key] = saved
                        item['_resolved_path'] = saved
                # Re-encode with resolved paths for the utility
                signatures_json = _json.dumps(placements)

            input_path = save_uploaded_file(uploaded_file)
            output_path = sign_pdf(
                input_path, uploaded_file.name,
                signature_image_path=sig_image_path,
                signature_data=signature_data,
                page_number=page_number,
                x=sig_x, y=sig_y,
                width=sig_width, height=sig_height,
                signatures_json=signatures_json if signatures_json else None,
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
            # Clean up any resolved signature files
            for fp in sig_file_paths.values():
                try:
                    os.remove(fp)
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

    # ── Video Converter ──
    if tool_slug == 'video-converter':
        if 'file' not in request.FILES:
            return JsonResponse({'error': 'No file was uploaded.'}, status=400)

        uploaded_file = request.FILES['file']
        output_format = request.POST.get('output_format', 'mp4')

        try:
            input_path = save_uploaded_file(uploaded_file)
            output_path = convert_video_format(
                input_path, uploaded_file.name, output_format
            )
            # Force application/octet-stream to prevent browser extensions (like IDM) 
            # from intercepting video/mp4 responses and aborting the fetch stream.
            content_type = 'application/octet-stream'

            try:
                os.remove(input_path)
            except OSError:
                pass
                
            from .utils import format_download_name
            final_filename = format_download_name(uploaded_file.name, output_path)

            import base64
            with open(output_path, 'rb') as f:
                b64_data = base64.b64encode(f.read()).decode('utf-8')

            try:
                os.remove(output_path)
            except OSError:
                pass

            return JsonResponse({
                'success': True,
                'filename': final_filename,
                'content_type': content_type,
                'data': b64_data
            })
        except Exception as e:
            return JsonResponse({'error': f'Video conversion failed: {str(e)}'}, status=500)

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

        return create_cleanup_response(output_path, content_type=content_type, filename=uploaded_file.name)

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
    
    import requests
    
    # In local development, determine the true public internet IP
    if ip in ['127.0.0.1', '::1', 'localhost'] or ip.startswith('192.168.') or ip.startswith('10.') or ip.startswith('172.16.'):
        try:
            ip = requests.get('https://api.ipify.org', timeout=3).text.strip()
        except:
            pass

    client_data = {
        'ip': ip if ip else 'Unavailable',
        'city': 'Unavailable',
        'country_name': '',
        'org': 'Unavailable'
    }
    
    if client_data['ip'] != 'Unavailable':
        try:
            # Determine IP Type
            client_data['ip_type'] = 'IPv6' if ':' in client_data['ip'] else 'IPv4'
            
            # Using ip-api.com which has higher reliability without User-Agent blocks
            resp = requests.get(f"http://ip-api.com/json/{ip}", timeout=4)
            if resp.status_code == 200:
                data = resp.json()
                if data.get('status') == 'success':
                    client_data['city'] = data.get('city', 'Unavailable')
                    client_data['region'] = data.get('regionName', 'Unavailable')
                    client_data['country'] = data.get('country', 'Unavailable')
                    client_data['countryCode'] = data.get('countryCode', 'Unavailable')
                    client_data['timezone'] = data.get('timezone', 'Unavailable')
                    client_data['lat'] = data.get('lat', 'Unavailable')
                    client_data['lon'] = data.get('lon', 'Unavailable')
                    
                    region = data.get('regionName', '')
                    country = data.get('country', '')
                    loc_parts = [p for p in [client_data['city'], region, country] if p and p != 'Unavailable']
                    if loc_parts:
                        client_data['full_location'] = ', '.join(loc_parts)
                    else:
                        client_data['full_location'] = 'Unavailable'
                        
                    client_data['isp'] = data.get('isp', 'Unavailable')
                    client_data['org'] = data.get('org', 'Unavailable')
                    client_data['asn'] = data.get('as', 'Unavailable')
        except:
            client_data['full_location'] = 'Unavailable'
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


@csrf_exempt
def currency_rates(request):
    """
    Optional privacy-safe live currency rates endpoint.
    No user conversion data is accepted or stored.
    """
    supported = ['INR', 'USD', 'EUR', 'GBP', 'AED', 'JPY', 'SGD', 'AUD', 'CAD', 'CNY']
    base = (request.GET.get('base') or 'USD').upper()
    if base not in supported:
        return JsonResponse({'error': 'Unsupported base currency.'}, status=400)

    now = time.time()
    # Runtime memory cache only (not persisted on disk/db)
    if (
        _CURRENCY_CACHE['base'] == base
        and _CURRENCY_CACHE['rates']
        and (now - _CURRENCY_CACHE['updated_at']) < 300
    ):
        return JsonResponse({
            'mode': 'live',
            'base': base,
            'rates': _CURRENCY_CACHE['rates'],
            'last_updated': _CURRENCY_CACHE['updated_at'],
        })

    api_key = os.environ.get('EXCHANGE_RATE_API_KEY')
    api_url = os.environ.get('EXCHANGE_RATE_API_URL')
    # If env vars are not configured, client should use fallback rates.
    if not api_key or not api_url:
        return JsonResponse({
            'mode': 'fallback',
            'base': base,
            'message': 'Live API not configured.',
        }, status=503)

    try:
        query = urllib.parse.urlencode({'base': base, 'symbols': ','.join(supported)})
        url = f"{api_url}?{query}"
        req = urllib.request.Request(url, headers={'apikey': api_key})
        with urllib.request.urlopen(req, timeout=6) as resp:
            payload = json.loads(resp.read().decode('utf-8'))

        rates = payload.get('rates') or {}
        filtered = {code: float(rates[code]) for code in supported if code in rates}
        filtered[base] = 1.0
        if len(filtered) < 2:
            raise ValueError('Insufficient rates from live API.')

        _CURRENCY_CACHE['base'] = base
        _CURRENCY_CACHE['rates'] = filtered
        _CURRENCY_CACHE['updated_at'] = now
        return JsonResponse({
            'mode': 'live',
            'base': base,
            'rates': filtered,
            'last_updated': now,
        })
    except Exception:
        return JsonResponse({
            'mode': 'fallback',
            'base': base,
            'message': 'Live rates unavailable.',
        }, status=503)
