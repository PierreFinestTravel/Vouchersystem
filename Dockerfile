# Dockerfile for Finest Travel Africa Voucher System
# Lightweight version - NO PDF conversion (returns DOCX files)

FROM python:3.11-slim

# Set environment variables
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
ENV PORT=10000

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

# Expose the port
EXPOSE 10000

# Run the application
CMD uvicorn app.main:app --host 0.0.0.0 --port ${PORT:-10000}

