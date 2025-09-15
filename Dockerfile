# Use official Python 3.12 slim image
FROM python:3.12.2-slim

# Set working directory inside the container
WORKDIR /app

# Prevent Python from writing pyc files and buffer stdout/stderr
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

# Install system dependencies needed for pandas, lxml, and other packages
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    build-essential \
    gcc \
    g++ \
    libffi-dev \
    libxml2-dev \
    libxslt1-dev \
    zlib1g-dev \
    libbz2-dev \
    liblzma-dev \
    curl \
    git \
    && rm -rf /var/lib/apt/lists/*

# Copy only requirements first to leverage Docker cache
COPY requirements.txt .

# Upgrade pip and install Python dependencies
RUN python -m pip install --upgrade pip setuptools wheel
RUN pip install --no-cache-dir -r requirements.txt

# Copy the entire project into the container
COPY . .

# Expose Dash default port
EXPOSE 8050

# Set environment variables for Flask/Dash (optional)
ENV FLASK_ENV=production
ENV DASH_DEBUG=False

# Command to run your app
CMD ["python", "extract_data_from_screener_site.py"]
