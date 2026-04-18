import os
import sys

# Set up Django environment
sys.path.append(os.getcwd())
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'All_In_One_PDF.settings')
import django
django.setup()

try:
    from converter.utils import generate_qr_code
    path = generate_qr_code("Test QR", output_format='png')
    print(f"SUCCESS: {path}")
    if os.path.exists(path):
        print(f"File exists: {path}")
        os.remove(path)
except Exception as e:
    import traceback
    print(f"FAILURE: {e}")
    traceback.print_exc()
