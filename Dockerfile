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

# Create upload directory
RUN mkdir -p /data/uploads

EXPOSE 5001

# Production: use gunicorn
CMD ["gunicorn", "--bind", "0.0.0.0:5001", "--workers", "2", "--timeout", "300", "app:app"]
