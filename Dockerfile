FROM python:3.11-slim

WORKDIR /app

# Install system deps for .doc file support.
# libgl1 is required by opencv-python (pulled in by rapidocr_onnxruntime) for
# the scanned-PDF OCR fallback; fonts-noto-cjk improves OCR glyph coverage.
RUN apt-get update && apt-get install -y --no-install-recommends \
    antiword \
    libgl1 \
    fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/*

# Install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy app code
COPY app.py .
COPY templates/ templates/
COPY static/ static/

# Create upload + history directories. UPLOAD_FOLDER / HISTORY_DIR point at
# them so uploads survive worker restarts and analysis history persists
# across container recreations (mount both as volumes in docker-compose).
RUN mkdir -p /data/uploads /data/history
ENV UPLOAD_FOLDER=/data/uploads
ENV HISTORY_DIR=/data/history
# 上传不设上限（默认）；OCR 完全放开（0=不限）。
# 可按需覆盖：MAX_CONTENT_LENGTH_MB / OCR_TIME_BUDGET / OCR_MAX_PAGES
ENV ANALYSIS_TIMEOUT=3600

EXPOSE 5001

# Production: use gunicorn. --timeout must exceed ANALYSIS_TIMEOUT (3600)
# so the analysis thread can flush a timeout error before the worker kills it.
CMD ["gunicorn", "--bind", "0.0.0.0:5001", "--workers", "2", "--timeout", "3900", "app:app"]
