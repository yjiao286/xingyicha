FROM python:3.11-slim

WORKDIR /app

# Install system deps for .doc file support
RUN apt-get update && apt-get install -y --no-install-recommends \
    antiword \
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
ENV MAX_CONTENT_LENGTH_MB=200
ENV ANALYSIS_TIMEOUT=270

EXPOSE 5001

# Production: use gunicorn
CMD ["gunicorn", "--bind", "0.0.0.0:5001", "--workers", "2", "--timeout", "300", "app:app"]
