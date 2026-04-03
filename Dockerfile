# Dockerfile
FROM python:3.9-slim

# Install OS-level deps needed to build Pillow (pdfplumber's imaging backend),
# plus general build tools for any Python C extensions
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
      build-essential \
      python3-dev \
      libxml2-dev \
      libxslt1-dev \
      libjpeg-dev \
      zlib1g-dev && \
    rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

RUN mkdir -p /app/data /app/failed_uploads

ENV PORT=8000
EXPOSE 8000

CMD ["gunicorn", "--bind", "0.0.0.0:8000", "--workers", "2", "--timeout", "120", "app:app"]
