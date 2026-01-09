# Dockerfile for Finest Travel Africa Voucher System
# Optimized for Render.com free tier with LibreOffice for PDF conversion

FROM python:3.11-slim

# Set environment variables
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
ENV PORT=10000

# Install system dependencies including LibreOffice for PDF conversion
# Using minimal LibreOffice installation to reduce image size
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer-nogui \
    fonts-liberation \
    fonts-dejavu-core \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/* \
    && rm -rf /var/cache/apt/*

# Set working directory
WORKDIR /app

# Copy requirements first (for better Docker layer caching)
COPY requirements.txt .

# Install Python dependencies
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY . .

# Create necessary directories
RUN mkdir -p /app/templates /tmp/voucher_gen

# Expose the port (Render uses PORT env variable)
EXPOSE 10000

# Run the application - Render sets PORT automatically
CMD uvicorn app.main:app --host 0.0.0.0 --port ${PORT:-10000}

