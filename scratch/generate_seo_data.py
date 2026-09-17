import os
import json

raw_keywords = {
    'word-to-pdf': ('Word to PDF Converter', 'Convert Word documents to PDF online for free with fast, secure, and high-quality conversion.'),
    'pptx-to-pdf': ('PPT to PDF', 'Convert PowerPoint presentations to PDF while preserving formatting and layout.'),
    'excel-to-pdf': ('Excel to PDF', 'Convert Excel spreadsheets to PDF quickly without losing formatting.'),
    'html-to-pdf': ('HTML to PDF Converter', 'Convert HTML pages or code into high-quality PDF documents online.'),
    'pdf-to-image': ('PDF to JPG', 'Convert PDF pages into JPG or PNG images instantly with excellent quality.'),
    'pdf-to-word': ('PDF to Word', 'Convert PDF files into editable Word documents accurately and quickly.'),
    'pdf-to-pptx': ('PDF to PPT', 'Convert PDF files into editable PowerPoint presentations online.'),
    'pdf-to-excel': ('PDF to Excel', 'Extract tables from PDF and convert them into editable Excel spreadsheets.'),
    'merge-word': ('Merge Word Documents', 'Combine multiple Word documents into one file without losing formatting.'),
    'png-to-jpg': ('PNG to JPG Converter', 'Convert PNG images into JPG format with optimized file size and quality.'),
    'jpg-to-png': ('JPG to PNG Converter', 'Convert JPG images into PNG format with transparent background support.'),
    'html-to-image': ('HTML to Image', 'Convert HTML webpages into PNG or JPG images instantly.'),
    'image-to-pdf': ('Image to PDF', 'Convert JPG, PNG, and other images into high-quality PDF files online.'),
    'pdf-to-pdfa': ('PDF to PDF/A', 'Convert PDF documents into PDF/A format for long-term archiving.'),
    'merge-pdf': ('Merge PDF Online', 'Combine multiple PDF files into one document quickly and securely.'),
    'split-pdf': ('Split PDF', 'Split large PDF files into separate pages or custom page ranges.'),
    'compress-pdf': ('Compress PDF', 'Reduce PDF file size without compromising document quality.'),
    'remove-pages': ('Remove PDF Pages', 'Delete unwanted pages from your PDF document in seconds.'),
    'extract-pages': ('Extract PDF Pages', 'Extract selected pages from PDF files into a new document.'),
    'organize-pdf': ('Organize PDF', 'Rearrange, reorder, or organize PDF pages easily online.'),
    'repair-pdf': ('Repair PDF', 'Repair corrupted PDF files and recover damaged documents.'),
    'ocr-pdf': ('OCR PDF', 'Convert scanned PDFs into searchable and editable documents using OCR.'),
    'rotate-pdf': ('Rotate PDF', 'Rotate PDF pages to the correct orientation online.'),
    'add-watermark': ('Add Watermark to PDF', 'Add text or image watermarks to PDF files for branding and security.'),
    'remove-watermark': ('Remove Watermark PDF', 'Remove unwanted watermarks from PDF documents quickly.'),
    'crop-pdf': ('Crop PDF', 'Crop PDF pages to remove unwanted margins and content.'),
    'edit-pdf': ('Edit PDF Online', 'Edit text, images, and pages in your PDF document easily.'),
    'unlock-pdf': ('Unlock PDF', 'Remove password protection from PDF files securely.'),
    'protect-pdf': ('Password Protect PDF', 'Encrypt PDF files with a secure password to protect sensitive information.'),
    'sign-pdf': ('Sign PDF Online', 'Add digital signatures or handwritten signatures to PDF documents.'),
    'redact-pdf': ('Redact PDF', 'Permanently remove confidential information from PDF documents.'),
    'scale-image': ('Scale Image', 'Scale images without losing quality using our online image scaler.'),
    'crop-image': ('Crop Image', 'Crop images to any size or aspect ratio quickly and easily.'),
    'resize-image': ('Resize Image', 'Resize JPG, PNG, WEBP, and GIF images online for free.'),
    'rotate-image': ('Rotate Image', 'Rotate images to any angle with high-quality output.'),
    'watermark-image': ('Watermark Image', 'Add text or image watermarks to your photos online.'),
    'compress-image': ('Compress Image', 'Compress JPG, PNG, and WEBP images while maintaining quality.'),
    'blur-image': ('Blur Image', 'Blur faces, backgrounds, or selected areas of an image.'),
    'brighten-image': ('Brighten Image', 'Adjust image brightness and improve photo visibility instantly.'),
    'change-gif-speed': ('GIF Speed Changer', 'Increase or decrease GIF animation speed online.'),
    'change-background': ('Change Image Background', 'Replace or change image backgrounds in seconds.'),
    'cut-image': ('Cut Image', 'Cut or trim unwanted parts of an image online.'),
    'merge-images': ('Merge Images', 'Combine multiple images into a single image easily.'),
    'remove-background': ('Remove Background', 'Remove image backgrounds automatically using AI.'),
    'image-converter': ('Image Converter', 'Convert images between JPG, PNG, WEBP, GIF, TIFF, BMP, and more.'),
    'jpg-converter': ('JPG Converter', 'Convert images to and from JPG format quickly online.'),
    'png-converter': ('PNG Converter', 'Convert images to PNG while maintaining transparency.'),
    'bmp-converter': ('BMP Converter', 'Convert BMP images into modern image formats.'),
    'gif-converter': ('GIF Converter', 'Convert GIF images into JPG, PNG, WEBP, or PDF.'),
    'pdf-converter': ('PDF Converter', 'Convert PDF documents into multiple supported formats.'),
    'tiff-converter': ('TIFF Converter', 'Convert TIFF images into JPG, PNG, WEBP, and PDF.'),
    'webp-converter': ('WEBP Converter', 'Convert WEBP images into JPG, PNG, GIF, or PDF.'),
    'dng-converter': ('DNG Converter', 'Convert DNG RAW images into common image formats.'),
    'password-generator': ('Password Generator', 'Generate strong, secure, and random passwords instantly.'),
    'qrcode-generator': ('QR Code Generator', 'Create free QR codes for URLs, text, email, Wi-Fi, and more.'),
    'meme-generator': ('Meme Generator', 'Create funny memes online using customizable templates.'),
    'name-generator': ('Name Generator', 'Generate unique business, brand, company, or baby names instantly.'),
    'chemical-balancer': ('Chemical Equation Balancer', 'Balance chemical equations accurately with an online calculator.'),
    'smart-creators': ('AI Content Creator', 'Create text, images, and digital assets using smart AI-powered tools.'),
    'story-generator': ('AI Story Generator', 'Generate creative stories, novels, scripts, and essays using AI in seconds.'),
    'unit-converter': ('Unit Converter', 'Convert length, weight, temperature, area, volume, speed, and more instantly.'),
    'speed-test': ('Internet Speed Test', 'Test your internet download, upload, and ping speed online for free.'),
    'audio-editor': ('Audio Editor Online', 'Edit audio files by trimming, cutting, adjusting volume, and applying effects.'),
    'merge-audio': ('Merge Audio Files', 'Combine multiple audio files into a single track online.'),
    'extract-audio': ('Extract Audio from Video', 'Extract MP3 or WAV audio from video files quickly and easily.')
}

def generate_seo_dict():
    seo_data = {}
    
    for slug, (keyword, ui_desc) in raw_keywords.items():
        is_pdf = 'pdf' in slug.lower()
        is_image = 'image' in slug.lower() or slug.endswith('-converter') or slug in ['png-to-jpg', 'jpg-to-png', 'scale-image', 'crop-image', 'resize-image', 'rotate-image', 'watermark-image', 'compress-image', 'blur-image', 'brighten-image', 'change-gif-speed', 'change-background', 'cut-image', 'merge-images', 'remove-background']
        is_converter = 'to' in slug.lower() or 'converter' in slug.lower()
        is_generator = 'generator' in slug.lower()
        
        # Unique Steps based on type
        if 'to' in slug and 'pdf' in slug:
            steps = [
                {'title': f'Upload your file', 'description': f'Select the document you want to convert into a PDF.'},
                {'title': 'Start Conversion', 'description': f'Our system will process your file and safely convert it to PDF format.'},
                {'title': 'Download PDF', 'description': 'Save your new PDF document directly to your device.'}
            ]
        elif 'to' in slug and 'pdf' not in slug:
            steps = [
                {'title': 'Upload your file', 'description': f'Select the file you need to convert.'},
                {'title': 'Convert Format', 'description': f'The tool will transform the file to the new target format securely.'},
                {'title': 'Save Result', 'description': 'Download the converted file once processing is complete.'}
            ]
        elif 'merge' in slug:
            steps = [
                {'title': 'Select multiple files', 'description': f'Upload all the files you want to combine together.'},
                {'title': 'Organize & Merge', 'description': f'Arrange them in your preferred order and click merge.'},
                {'title': 'Download Merged File', 'description': 'Save the combined file instantly.'}
            ]
        elif is_image and 'compress' in slug or 'resize' in slug or 'crop' in slug:
             steps = [
                {'title': 'Upload image', 'description': f'Select the photo or image you want to optimize.'},
                {'title': 'Adjust Settings', 'description': f'Set your desired parameters and let the tool apply the changes.'},
                {'title': 'Download Image', 'description': 'Save your optimized image.'}
            ]
        else:
            steps = [
                {'title': 'Upload or Enter Data', 'description': f'Provide the required input for the {keyword} tool.'},
                {'title': 'Process', 'description': f'Our servers will execute the requested action.'},
                {'title': 'Download', 'description': 'Instantly retrieve your final result.'}
            ]
            
        SPECIFIC_FEATURES = {
            'word-to-pdf': [
                {'title': 'Word Document Conversion', 'description': 'Convert supported Word documents into PDF format.'},
                {'title': 'PDF Output', 'description': 'Generate a PDF version of the uploaded Word document.'},
                {'title': 'Document Layout', 'description': 'Retain the document structure as supported by the current conversion engine.'},
                {'title': 'Easy Download', 'description': 'Download the generated PDF after successful conversion.'}
            ],
            'pptx-to-pdf': [
                {'title': 'Presentation to PDF', 'description': 'Convert supported PowerPoint presentations into PDF documents.'},
                {'title': 'Slide Layout', 'description': 'Maintain slide structure as supported by the conversion engine.'},
                {'title': 'PDF Presentation Output', 'description': 'Create a PDF containing the converted presentation pages.'},
                {'title': 'Easy PDF Download', 'description': 'Download the completed PDF after conversion.'}
            ],
            'excel-to-pdf': [
                {'title': 'Spreadsheet to PDF', 'description': 'Convert supported spreadsheet files into PDF format.'},
                {'title': 'Worksheet Conversion', 'description': 'Turn spreadsheet content into a shareable PDF document.'},
                {'title': 'Table Layout', 'description': 'Preserve spreadsheet structure as supported by the converter.'},
                {'title': 'Downloadable PDF', 'description': 'Download the converted spreadsheet as a PDF.'}
            ],
            'html-to-pdf': [
                {'title': 'Webpage Conversion', 'description': 'Convert active HTML URLs or uploaded files into PDF.'},
                {'title': 'Capture Layouts', 'description': 'Render HTML structure and styling into the PDF document.'},
                {'title': 'Single PDF Output', 'description': 'Generate a readable PDF document from webpage content.'},
                {'title': 'Quick Download', 'description': 'Save the rendered HTML-to-PDF file securely.'}
            ],
            'pdf-to-image': [
                {'title': 'PDF to JPG/PNG', 'description': 'Convert individual PDF pages into high-resolution image files.'},
                {'title': 'Page Extraction', 'description': 'Process every page of the PDF document into a separate image.'},
                {'title': 'Visual Accuracy', 'description': 'Retain the visual content of the PDF in the image output.'},
                {'title': 'Image Download', 'description': 'Download the converted image files as a ZIP archive.'}
            ],
            'pdf-to-word': [
                {'title': 'Editable Output', 'description': 'Convert PDF files back into editable Microsoft Word format.'},
                {'title': 'Text Extraction', 'description': 'Extract text and paragraphs from the source PDF document.'},
                {'title': 'Structure Preservation', 'description': 'Keep the formatting as closely aligned to the PDF as possible.'},
                {'title': 'DOCX Download', 'description': 'Download the resulting DOCX file for immediate editing.'}
            ],
            'pdf-to-pptx': [
                {'title': 'Presentation Recovery', 'description': 'Convert PDF files into editable PowerPoint slides.'},
                {'title': 'Slide Mapping', 'description': 'Map individual PDF pages to corresponding presentation slides.'},
                {'title': 'Content Extraction', 'description': 'Extract text and layouts for presentation editing.'},
                {'title': 'PPTX Download', 'description': 'Save the converted PowerPoint file.'}
            ],
            'pdf-to-excel': [
                {'title': 'Data Extraction', 'description': 'Convert PDF tables into editable Excel spreadsheet formats.'},
                {'title': 'Table Recognition', 'description': 'Identify tabular data structures within the PDF document.'},
                {'title': 'Spreadsheet Output', 'description': 'Format the extracted data into XLSX format.'},
                {'title': 'Excel Download', 'description': 'Download the resulting spreadsheet for data analysis.'}
            ],
            'merge-pdf': [
                {'title': 'Combine Multiple PDFs', 'description': 'Join multiple PDF documents into a single PDF.'},
                {'title': 'Control File Order', 'description': 'Arrange uploaded PDFs in the required sequence before merging.'},
                {'title': 'Single PDF Output', 'description': 'Create one combined PDF from multiple source documents.'},
                {'title': 'Simplified Document Management', 'description': 'Combine related PDF documents for easier storage and sharing.'}
            ],
            'split-pdf': [
                {'title': 'Split PDF by Page', 'description': 'Separate a multi-page PDF into individual pages or sections.'},
                {'title': 'Extract Required Pages', 'description': 'Keep only the pages you need without manually recreating the document.'},
                {'title': 'Create Separate PDF Files', 'description': 'Divide a large PDF into smaller PDF documents.'},
                {'title': 'Simple Page Selection', 'description': 'Choose the required pages through the existing Split PDF workflow.'}
            ],
            'compress-pdf': [
                {'title': 'Reduce PDF File Size', 'description': 'Compress PDF documents to create smaller files.'},
                {'title': 'Maintain Usable Quality', 'description': 'Reduce file size while keeping document quality suitable for normal use.'},
                {'title': 'Easier File Sharing', 'description': 'Create smaller PDFs that are easier to upload or share.'},
                {'title': 'Compressed PDF Download', 'description': 'Download the optimized PDF after processing.'}
            ],
            'rotate-pdf': [
                {'title': 'Rotate PDF Pages', 'description': 'Change page orientation using the available rotation controls.'},
                {'title': 'Correct Page Orientation', 'description': 'Fix incorrectly oriented pages in a PDF.'},
                {'title': 'Page Rotation Control', 'description': 'Apply supported rotation options to the selected PDF pages.'},
                {'title': 'Download Rotated PDF', 'description': 'Save the corrected PDF after processing.'}
            ],
            'remove-pages': [
                {'title': 'Remove Unwanted Pages', 'description': 'Delete pages that are no longer required.'},
                {'title': 'Select PDF Pages', 'description': 'Choose the pages to remove using the current interface.'},
                {'title': 'Keep Remaining Content', 'description': 'Create a new PDF containing the pages that remain.'},
                {'title': 'Download Updated PDF', 'description': 'Save the processed document after page removal.'}
            ],
            'extract-pages': [
                {'title': 'Extract Selected Pages', 'description': 'Choose specific pages from an existing PDF.'},
                {'title': 'Create a New PDF', 'description': 'Generate a separate PDF from the selected pages.'},
                {'title': 'Flexible Page Selection', 'description': 'Select individual pages or ranges.'},
                {'title': 'Quick PDF Download', 'description': 'Download the extracted document after processing.'}
            ],
            'organize-pdf': [
                {'title': 'Reorder PDF Pages', 'description': 'Arrange pages into the required sequence.'},
                {'title': 'Organize Document Structure', 'description': 'Correct the page order of an existing PDF.'},
                {'title': 'Visual Page Management', 'description': 'Use page previews for organization.'},
                {'title': 'Save Organized PDF', 'description': 'Download the reordered document.'}
            ],
            'repair-pdf': [
                {'title': 'Document Recovery', 'description': 'Attempt to recover content from corrupted PDF files.'},
                {'title': 'Structure Repair', 'description': 'Fix underlying PDF structure issues where possible.'},
                {'title': 'Salvage Data', 'description': 'Extract readable pages from damaged documents.'},
                {'title': 'Repaired PDF Download', 'description': 'Save the recovered document.'}
            ],
            'ocr-pdf': [
                {'title': 'Text Extraction', 'description': 'Extract text from scanned PDF documents using OCR.'},
                {'title': 'Searchable PDF', 'description': 'Create a searchable document from image-based PDFs.'},
                {'title': 'Character Recognition', 'description': 'Recognize printed text within images.'},
                {'title': 'Editable Output', 'description': 'Save the OCR-processed document for future editing.'}
            ],
            'add-watermark': [
                {'title': 'Apply Watermark', 'description': 'Add text or image overlays to PDF documents.'},
                {'title': 'Protect Content', 'description': 'Mark documents as draft, confidential, or copyrighted.'},
                {'title': 'Custom Positioning', 'description': 'Place the watermark on the desired pages.'},
                {'title': 'Watermarked PDF Download', 'description': 'Save the protected document.'}
            ],
            'remove-watermark': [
                {'title': 'Watermark Removal', 'description': 'Attempt to remove overlays from PDF documents.'},
                {'title': 'Clean Documents', 'description': 'Remove draft or sample marks from final files.'},
                {'title': 'Batch Processing', 'description': 'Process multiple pages simultaneously.'},
                {'title': 'Clean PDF Download', 'description': 'Save the document without the watermark.'}
            ],
            'crop-pdf': [
                {'title': 'Crop PDF Pages', 'description': 'Remove unwanted margins or content from PDF pages.'},
                {'title': 'Custom Dimensions', 'description': 'Adjust the visible page area.'},
                {'title': 'Uniform Cropping', 'description': 'Apply the same crop box to multiple pages.'},
                {'title': 'Cropped PDF Download', 'description': 'Save the resized document.'}
            ],
            'edit-pdf': [
                {'title': 'PDF Modification', 'description': 'Make basic modifications to PDF document content.'},
                {'title': 'Add Elements', 'description': 'Insert text, shapes, or annotations.'},
                {'title': 'Form Filling', 'description': 'Fill out interactive PDF forms.'},
                {'title': 'Edited PDF Download', 'description': 'Save your changes to a new file.'}
            ],
            'unlock-pdf': [
                {'title': 'Remove Passwords', 'description': 'Remove security restrictions from unlocked PDF files.'},
                {'title': 'Enable Editing', 'description': 'Remove restrictions on copying and printing.'},
                {'title': 'Streamline Access', 'description': 'Make documents easier to open without passwords.'},
                {'title': 'Unlocked PDF Download', 'description': 'Save the unrestricted document.'}
            ],
            'protect-pdf': [
                {'title': 'Password Protect PDF', 'description': 'Add password protection to supported PDF documents.'},
                {'title': 'Restrict Unauthorized Access', 'description': 'Protect the document using available security options.'},
                {'title': 'PDF Encryption', 'description': 'Apply the encryption method supported by the current backend.'},
                {'title': 'Protected PDF Download', 'description': 'Download the password-protected output.'}
            ],
            'sign-pdf': [
                {'title': 'Digital Signatures', 'description': 'Add visual signatures to PDF documents.'},
                {'title': 'Document Approval', 'description': 'Mark documents as approved or finalized.'},
                {'title': 'Draw Signature', 'description': 'Draw or upload a signature image.'},
                {'title': 'Signed PDF Download', 'description': 'Save the signed document securely.'}
            ],
            'redact-pdf': [
                {'title': 'Information Redaction', 'description': 'Black out sensitive text or images in PDF documents.'},
                {'title': 'Protect Privacy', 'description': 'Ensure confidential data is permanently obscured.'},
                {'title': 'Visual Verification', 'description': 'Preview redactions before finalizing.'},
                {'title': 'Redacted PDF Download', 'description': 'Save the secure document.'}
            ],
            'resize-image': [
                {'title': 'Custom Image Dimensions', 'description': 'Set the required width and height.'},
                {'title': 'Resize Supported Images', 'description': 'Change image dimensions using accepted formats.'},
                {'title': 'Maintain Suitable Quality', 'description': 'Resize images while retaining output quality.'},
                {'title': 'Download Resized Image', 'description': 'Save the resized result.'}
            ],
            'crop-image': [
                {'title': 'Custom Image Cropping', 'description': 'Crop an image to the required area.'},
                {'title': 'Remove Unwanted Areas', 'description': 'Keep the important portion of an image.'},
                {'title': 'Crop Dimensions', 'description': 'Adjust the crop area using the controls actually available.'},
                {'title': 'Download Cropped Image', 'description': 'Save the cropped result.'}
            ],
            'rotate-image': [
                {'title': 'Rotate Your Image', 'description': 'Change image orientation.'},
                {'title': 'Correct Image Orientation', 'description': 'Fix sideways or incorrectly oriented photos.'},
                {'title': 'Rotation Controls', 'description': 'Use the angles supported by the interface.'},
                {'title': 'Download Rotated Image', 'description': 'Save the corrected image.'}
            ],
            'compress-image': [
                {'title': 'Reduce Image File Size', 'description': 'Compress supported images into smaller files.'},
                {'title': 'Optimize Image Storage', 'description': 'Reduce the amount of storage required.'},
                {'title': 'Maintain Suitable Image Quality', 'description': 'Balance compression and output quality.'},
                {'title': 'Download Compressed Image', 'description': 'Save the optimized result.'}
            ],
            'blur-image': [
                {'title': 'Apply Image Blur', 'description': 'Add a blur effect using the current tool controls.'},
                {'title': 'Adjust Blur', 'description': 'Change blur intensity using available controls.'},
                {'title': 'Create Soft Image Effects', 'description': 'Use blur for visual effects or obscuring areas.'},
                {'title': 'Download Blurred Image', 'description': 'Save the processed image.'}
            ],
            'brighten-image': [
                {'title': 'Adjust Image Brightness', 'description': 'Increase or modify image brightness.'},
                {'title': 'Improve Dark Images', 'description': 'Make dark image content easier to see.'},
                {'title': 'Brightness Control', 'description': 'Adjust intensity according to available controls.'},
                {'title': 'Download Brightened Image', 'description': 'Save the processed image.'}
            ],
            'change-gif-speed': [
                {'title': 'Adjust GIF Animation Speed', 'description': 'Change how quickly the animated GIF plays.'},
                {'title': 'Speed Up GIF', 'description': 'Increase animation playback speed.'},
                {'title': 'Slow Down GIF', 'description': 'Reduce animation playback speed.'},
                {'title': 'Download Updated GIF', 'description': 'Save the GIF with the selected speed.'}
            ],
            'change-background': [
                {'title': 'Change Image Background', 'description': 'Replace the existing background.'},
                {'title': 'Choose a New Background', 'description': 'Use the background options actually available.'},
                {'title': 'Preview Background Changes', 'description': 'Show a preview before final download.'},
                {'title': 'Download Updated Image', 'description': 'Save the final image.'}
            ],
            'remove-background': [
                {'title': 'Remove Image Background', 'description': 'Separate the main subject from its background.'},
                {'title': 'Transparent Background', 'description': 'Create transparency in the output image.'},
                {'title': 'Subject Extraction', 'description': 'Keep the foreground subject while removing surroundings.'},
                {'title': 'Download Processed Image', 'description': 'Save the background-removed result.'}
            ],
            'watermark-image': [
                {'title': 'Add Image Watermark', 'description': 'Overlay text or logos onto your image.'},
                {'title': 'Protect Copyright', 'description': 'Mark your photos to prevent unauthorized use.'},
                {'title': 'Custom Placement', 'description': 'Position the watermark appropriately.'},
                {'title': 'Watermarked Image Download', 'description': 'Save the protected photo.'}
            ],
            'scale-image': [
                {'title': 'Image Scaling', 'description': 'Scale images up or down proportionately.'},
                {'title': 'Preserve Aspect Ratio', 'description': 'Maintain the original width-to-height ratio.'},
                {'title': 'Quality Retention', 'description': 'Scale using optimized resampling algorithms.'},
                {'title': 'Scaled Image Download', 'description': 'Save the resized image file.'}
            ],
            'merge-word': [
                {'title': 'Combine Documents', 'description': 'Merge multiple Word files into one.'},
                {'title': 'Sequential Assembly', 'description': 'Order documents before combining them.'},
                {'title': 'Unified Formatting', 'description': 'Output a single contiguous DOCX file.'},
                {'title': 'Merged DOCX Download', 'description': 'Download the combined document.'}
            ],
            'image-converter': [
                {'title': 'Universal Conversion', 'description': 'Convert images between popular formats.'},
                {'title': 'Format Flexibility', 'description': 'Support for JPG, PNG, BMP, WEBP, and TIFF.'},
                {'title': 'Quality Preservation', 'description': 'Retain color depth and resolution.'},
                {'title': 'Converted Image Download', 'description': 'Save the image in the new format.'}
            ],
            'jpg-converter': [
                {'title': 'Convert to JPG', 'description': 'Convert PNG, BMP, and WEBP to JPG.'},
                {'title': 'Optimize Size', 'description': 'Utilize JPG compression for smaller file sizes.'},
                {'title': 'Broad Compatibility', 'description': 'Ensure the image works on all platforms.'},
                {'title': 'JPG Download', 'description': 'Download the converted JPG file.'}
            ],
            'png-converter': [
                {'title': 'Convert to PNG', 'description': 'Convert JPG, BMP, and WEBP to PNG.'},
                {'title': 'Transparency Support', 'description': 'Maintain alpha channels and transparency.'},
                {'title': 'Lossless Quality', 'description': 'Save images without compression artifacts.'},
                {'title': 'PNG Download', 'description': 'Download the converted PNG file.'}
            ],
            'image-to-pdf': [
                {'title': 'Images to PDF', 'description': 'Convert JPG, PNG, and other images into PDF.'},
                {'title': 'Multiple Images', 'description': 'Combine multiple images into a single PDF document.'},
                {'title': 'Document Compilation', 'description': 'Create scanned-style documents from photos.'},
                {'title': 'PDF Download', 'description': 'Save the compiled PDF document.'}
            ]
        }
        
        # Fallback generator for remaining tools
        if slug in SPECIFIC_FEATURES:
            features = SPECIFIC_FEATURES[slug]
        else:
            features = [
                {'title': f'{keyword} Processing', 'description': f'Process your data using the {keyword} engine.'},
                {'title': 'Accurate Execution', 'description': f'Perform the {keyword} operation reliably.'},
                {'title': 'Rapid Output', 'description': 'Generate the output quickly without delays.'},
                {'title': 'Secure Download', 'description': 'Download the finalized file or result.'}
            ]
            
        # Unique FAQs based on specific slug
        faqs = []
        if slug == 'word-to-pdf':
            faqs = [
                {'question': 'How do I convert Word to PDF without losing formatting?', 'answer': 'Simply upload your DOC or DOCX file to our tool. We use advanced conversion engines that preserve your original layout, fonts, and images.'},
                {'question': 'Is this Word to PDF converter free?', 'answer': 'Yes, you can convert your Word documents to PDF for free directly in your browser.'},
                {'question': 'Can I convert DOCX to PDF?', 'answer': 'Absolutely! Our tool supports both older DOC formats and newer DOCX formats.'},
                {'question': 'Are my documents safe?', 'answer': 'Yes. We prioritize your privacy and ensure files are processed securely and not kept permanently.'},
                {'question': 'Do I need Microsoft Word installed?', 'answer': 'No, the conversion is handled entirely on our servers, so you don\'t need any software installed.'}
            ]
        elif slug == 'pptx-to-pdf':
            faqs = [
                {'question': 'How do I convert PowerPoint to PDF online?', 'answer': 'Upload your PowerPoint presentation, start the conversion, and download the generated PDF when processing is complete.'},
                {'question': 'Can I convert PPTX to PDF?', 'answer': 'Yes. The PowerPoint to PDF tool fully supports modern PPTX formats.'},
                {'question': 'Does PPT to PDF preserve slide formatting?', 'answer': 'Yes, the converter is designed to retain your presentation layout, fonts, and images.'},
                {'question': 'Do I need to install PowerPoint?', 'answer': 'No, our tool processes everything online without requiring Microsoft Office.'},
                {'question': 'How do I download the converted PDF?', 'answer': 'After successful conversion, a download button will appear allowing you to save the PDF.'}
            ]
        elif slug == 'excel-to-pdf':
            faqs = [
                {'question': 'How do I convert Excel to PDF?', 'answer': 'Upload your XLS or XLSX file, click convert, and download the resulting PDF.'},
                {'question': 'Will my spreadsheet fit on a single page?', 'answer': 'The converter automatically formats your spreadsheets to fit appropriately on standard PDF pages.'},
                {'question': 'Does this work without Microsoft Excel?', 'answer': 'Yes, you do not need Excel installed to use this online converter.'},
                {'question': 'Is my financial data secure?', 'answer': 'Yes, we take security seriously and files are deleted after processing.'},
                {'question': 'Can I convert multiple sheets?', 'answer': 'Yes, the tool typically converts the active data across your worksheets.'}
            ]
        elif slug == 'merge-pdf':
            faqs = [
                {'question': 'How do I combine PDF files?', 'answer': 'Upload multiple PDF files, arrange them in your preferred order, and click merge to combine them into one document.'},
                {'question': 'Is this PDF merger free?', 'answer': 'Yes, you can merge your PDF files together at no cost.'},
                {'question': 'Can I change the order of the PDFs?', 'answer': 'Yes, most of our tools allow you to specify or reorder the files before merging.'},
                {'question': 'Will merging reduce the quality?', 'answer': 'No, the original quality of the individual PDFs is preserved in the merged file.'},
                {'question': 'Is there a limit on how many files I can merge?', 'answer': 'There may be a reasonable file size or file count limit to ensure stable processing.'}
            ]
        elif slug == 'compress-pdf':
            faqs = [
                {'question': 'How do I reduce PDF file size?', 'answer': 'Upload your PDF to the compressor tool, and it will automatically optimize images and fonts to reduce the overall size.'},
                {'question': 'Will compressing the PDF ruin the quality?', 'answer': 'We use intelligent compression that reduces file size while maintaining readability and visual quality.'},
                {'question': 'Is this tool useful for email attachments?', 'answer': 'Yes, reducing the PDF size makes it much easier to attach to emails and share online.'},
                {'question': 'Can I compress scanned PDFs?', 'answer': 'Yes, scanned PDFs often contain large images, making them perfect candidates for our compression tool.'},
                {'question': 'Do I need to pay to compress a PDF?', 'answer': 'No, our basic compression tool is free to use.'}
            ]
        elif slug == 'remove-background':
            faqs = [
                {'question': 'How do I remove the background from an image?', 'answer': 'Simply upload your image, and our system will automatically detect the subject and remove the background.'},
                {'question': 'What image formats are supported?', 'answer': 'We generally support common formats like JPG and PNG for background removal.'},
                {'question': 'Will the result have a transparent background?', 'answer': 'Yes, the output is typically a PNG file with a transparent background, perfect for editing.'},
                {'question': 'Do I need to manually trace the subject?', 'answer': 'No, the process is fully automated.'},
                {'question': 'Is the background remover free?', 'answer': 'Yes, you can remove backgrounds online for free.'}
            ]
        elif slug == 'qrcode-generator':
             faqs = [
                {'question': 'How do I create a QR code?', 'answer': 'Enter your URL or text into the input field, customize the colors or style if available, and click generate.'},
                {'question': 'Are the QR codes permanent?', 'answer': 'Yes, static QR codes generated here will work permanently as long as the destination URL remains active.'},
                {'question': 'Can I add a logo to my QR code?', 'answer': 'Yes, our advanced generator allows you to upload a logo to place in the center of the QR code.'},
                {'question': 'What formats can I download the QR code in?', 'answer': 'You can typically download the generated QR code in high-quality PNG, JPG, or SVG formats.'},
                {'question': 'Is the QR Code Generator free?', 'answer': 'Yes, creating standard QR codes is completely free.'}
            ]
        else:
            # Dynamic fallback that is STILL UNIQUE to the tool
            verb = "convert" if is_converter else "process"
            faqs = [
                {'question': f'How do I use the {keyword} online?', 'answer': f'Upload your file or enter the required information, start the {verb} process, and download the resulting {keyword} output securely.'},
                {'question': f'Is the {keyword} tool free?', 'answer': f'Yes, you can use our {keyword} online without any cost.'},
                {'question': f'Does this work on mobile devices?', 'answer': f'Yes, the {keyword} is fully responsive and works seamlessly on smartphones and tablets.'},
                {'question': 'Are my files secure?', 'answer': 'We take privacy seriously. Files are processed securely and are not stored permanently.'},
                {'question': f'Do I need to install any software?', 'answer': f'No, the {keyword} works entirely in your web browser, requiring no downloads or installations.'}
            ]
            
        # About
        about = f"<p>The <strong>{keyword}</strong> is a specialized online tool built to help you easily manage and process your files. {ui_desc} Whether you are a student, a professional, or just an everyday user, this tool simplifies complex tasks so you can get work done faster.</p>"
        if is_converter:
             about += f"<p>Converting your files has never been simpler. Just upload your document, wait a few seconds for our secure servers to process the request, and download your high-quality output immediately.</p>"
        elif is_pdf:
             about += f"<p>Managing PDFs shouldn't require expensive desktop software. Our tool provides a straightforward, secure way to handle your PDF documents entirely online.</p>"
        elif is_image:
             about += f"<p>Editing and optimizing images is quick and easy. Our system ensures high-quality results while saving you time and effort.</p>"
             
        seo_data[slug] = {
            'seo_title': f"{keyword} – Free Online Tool | ScanPDF",
            'seo_description': ui_desc,
            'seo_h1': keyword,
            'steps': steps,
            'features': features,
            'cards': features, # reuse for benefit cards
            'faqs': faqs,
            'about': about
        }

    return seo_data

def main():
    data = generate_seo_dict()
    output_path = r"r:\DLK-Project\Scan-PDF\converter\seo_content.py"
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write("# Auto-generated SEO data\n")
        f.write("SEO_DATA = ")
        json.dump(data, f, indent=4)
        f.write("\n")

if __name__ == '__main__':
    main()
