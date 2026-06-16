# ScanPDF Video Processing — Deployment Guide

## Server Requirements (Hostinger VPS)

- Ubuntu 22.04 LTS
- 2+ CPU cores
- 4GB+ RAM (8GB recommended for 4K/2GB files)
- 50GB+ SSD storage

## Step 1: System Dependencies

```bash
sudo apt update && sudo apt upgrade -y
sudo apt install -y python3-pip python3-venv python3-dev nginx redis-server ffmpeg supervisor git
# For python-magic
sudo apt install -y libmagic1
```

## Step 2: Project Setup

```bash
sudo mkdir -p /var/www/scanpdf
sudo chown $USER:$USER /var/www/scanpdf
git clone <your-repo> /var/www/scanpdf
cd /var/www/scanpdf
python3 -m venv venv
source venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
```

## Step 3: FFmpeg Setup

```bash
# Verify FFmpeg is installed
ffmpeg -version
ffprobe -version
# If missing:
# sudo apt install ffmpeg
```

## Step 4: Environment & Static Files

```bash
cd /var/www/scanpdf
python manage.py collectstatic --noinput
python manage.py migrate
```

Create `.env`:
```
DEBUG=False
CELERY_BROKER_URL=redis://localhost:6379/0
CELERY_RESULT_BACKEND=redis://localhost:6379/0
```

## Step 5: Gunicorn + Systemd

```bash
sudo mkdir -p /run/scanpdf
sudo chown www-data:www-data /run/scanpdf
sudo cp deploy/gunicorn.service /etc/systemd/system/scanpdf-gunicorn.service
sudo cp deploy/celery.service /etc/systemd/system/scanpdf-celery.service
sudo systemctl daemon-reload
sudo systemctl enable scanpdf-gunicorn scanpdf-celery
sudo systemctl start scanpdf-gunicorn scanpdf-celery
```

## Step 6: Nginx

```bash
sudo cp deploy/nginx.conf /etc/nginx/sites-available/scanpdf
sudo ln -sf /etc/nginx/sites-available/scanpdf /etc/nginx/sites-enabled/
sudo rm -f /etc/nginx/sites-enabled/default
sudo nginx -t
sudo systemctl restart nginx
sudo systemctl enable nginx
```

## Step 7: Redis

```bash
sudo systemctl enable redis-server
sudo systemctl start redis-server
```

## Step 8: SSL (Certbot)

```bash
sudo apt install certbot python3-certbot-nginx
sudo certbot --nginx -d your-domain.com
```

## Monitoring

```bash
# Check services
sudo systemctl status scanpdf-gunicorn
sudo systemctl status scanpdf-celery
sudo systemctl status nginx
sudo systemctl status redis-server

# Check logs
sudo tail -f /var/log/nginx/error.log
sudo journalctl -u scanpdf-gunicorn -f
sudo journalctl -u scanpdf-celery -f
```

## Troubleshooting

**502 Bad Gateway:** Check gunicorn socket exists: `ls -la /run/scanpdf/gunicorn.sock`

**Large uploads failing:** Ensure `client_max_body_size 5G;` is in nginx.conf and restart nginx.

**FFmpeg not found:** Install with `sudo apt install ffmpeg` and verify `ffmpeg` is in PATH.

**Celery not processing:** Verify Redis is running: `redis-cli ping` should return `PONG`.

**Temp disk full:** Temp files auto-delete after 10 minutes. For manual cleanup: `python manage.py shell -c "from video_processor.cleanup import immediate_cleanup_all; immediate_cleanup_all('/tmp/scanpdf_video')"`

## Important Security Notes

- All uploaded media is stored in `/tmp/scanpdf_video/` and auto-deleted
- No user data or videos are persisted to the database
- Chunk uploads support files up to 2GB+ via 5MB chunks
- FFmpeg commands are sanitized against injection
- CSRF protection is active on all POST endpoints
