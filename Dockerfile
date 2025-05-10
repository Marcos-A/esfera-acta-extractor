# Dockerfile
FROM python:3.9-slim

# Install OS-level deps needed to build Pillow (pdfplumber’s imaging backend),
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

# Install Python libraries: pandas for data handling, pdfplumber for PDF table extraction
RUN pip install --no-cache-dir \
      pandas \
      pdfplumber

# Drop into a shell by default
CMD ["bash"]
