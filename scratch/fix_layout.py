import os
import glob

# 1. Update static error files
static_dir = r"r:\DLK-Scan-PDF\Scan-PDF\static\errors"
for filepath in glob.glob(os.path.join(static_dir, "*.html")):
    with open(filepath, "r", encoding="utf-8") as f:
        content = f.read()
    
    # Add flex layout to body
    content = content.replace(
        '<body class="bg-surface-50 text-surface-900 antialiased selection:bg-brand-500 selection:text-white">',
        '<body class="bg-surface-50 text-surface-900 antialiased selection:bg-brand-500 selection:text-white min-h-screen flex flex-col">'
    )
    # Fix section height
    content = content.replace(
        '<section class="min-h-[68vh] flex items-center justify-center relative overflow-hidden px-4">',
        '<section class="min-h-[calc(100vh-150px)] flex-grow flex items-center justify-center relative overflow-hidden px-4 py-12 w-full">'
    )
    
    with open(filepath, "w", encoding="utf-8") as f:
        f.write(content)

# 2. Update the base.html for Django errors again to add flex-grow and w-full
base_path = r"r:\DLK-Scan-PDF\Scan-PDF\templates\errors\base.html"
with open(base_path, "r", encoding="utf-8") as f:
    base_content = f.read()

base_content = base_content.replace(
    '<section class="min-h-[calc(100vh-150px)] flex items-center justify-center relative overflow-hidden px-4">',
    '<section class="min-h-[calc(100vh-150px)] flex-1 flex items-center justify-center relative overflow-hidden px-4 py-12 w-full">'
)

with open(base_path, "w", encoding="utf-8") as f:
    f.write(base_content)
