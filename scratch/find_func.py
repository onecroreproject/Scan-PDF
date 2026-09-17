import re

def main():
    try:
        with open(r'r:\DLK-Project\Scan-PDF\dynamic_qr\views.py', 'r', encoding='utf-8') as f:
            lines = f.readlines()
        for i, line in enumerate(lines):
            if 'def dqr_short_url_analytics_view' in line:
                print(f"Found on line {i+1}")
    except Exception as e:
        print("Error:", e)

if __name__ == '__main__':
    main()
