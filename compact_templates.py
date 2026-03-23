import os
import re

directory = r'r:\Balakrishnan\All_In_One_PDF\templates\converter'
files = [f for f in os.listdir(directory) if f.endswith('.html') and f not in ['base.html', 'home.html']]

replacements = [
    (r'p-6 sm:p-10 shadow-mega', 'p-4 sm:p-5 shadow-mega'),
    (r'p-10 sm:p-14 md:p-20', 'p-6 sm:p-8 md:p-10'),
    (r'mt-8', 'mt-4'),
    (r'mt-6', 'mt-3'),
    (r'mb-10 sm:mb-16', 'mb-4 sm:mb-6'),
    (r'pt-24 sm:pt-32 pb-12 sm:pb-20', 'pt-16 sm:pt-20 pb-8 sm:pb-12'),
    (r'py-8 px-6', 'py-4 px-5'),
    (r'w-20 h-20', 'w-14 h-14'),
    (r'w-12 h-12', 'w-9 h-9'),
    (r'w-10 h-10', 'w-8 h-8'),
    (r'text-3xl sm:text-4xl md:text-5xl lg:text-6xl', 'text-xl sm:text-2xl md:text-3xl'),
    (r'text-2xl font-black text-surface-900 mb-2', 'text-lg font-black text-surface-900 mb-1'),
    (r'text-2xl font-black text-surface-900 mb-3', 'text-lg font-black text-surface-900 mb-1'),
    (r'sm:py-5', 'sm:py-4'),
    (r'py-4 sm:py-5', 'py-3 sm:py-4'),
    (r'py-5', 'py-3'),
    (r'rounded-3xl', 'rounded-2xl'),
    (r'rounded-\[2\.5rem\]', 'rounded-2xl'),
    (r'rounded-\[2rem\]', 'rounded-xl'),
    (r'gap-8', 'gap-4'),
    (r'space-y-8', 'space-y-4'),
    (r'space-y-6', 'space-y-3'),
    (r'mb-5', 'mb-2'),
    (r'mb-8', 'mb-4'),
    (r'pt-8', 'pt-4'),
]

for filename in files:
    path = os.path.join(directory, filename)
    with open(path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    new_content = content
    for pattern, subt in replacements:
        new_content = re.sub(pattern, subt, new_content)
    
    if new_content != content:
        with open(path, 'w', encoding='utf-8') as f:
            f.write(new_content)
        print(f"Updated {filename}")
