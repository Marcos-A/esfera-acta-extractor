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

# Copy requirements file
COPY requirements.txt .

# Install all Python dependencies from requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

# Drop into a shell by default
CMD ["bash"]
