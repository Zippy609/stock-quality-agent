# Use official Python 3.12.2 slim image
FROM python:3.12.2-slim

# Set working directory inside the container
WORKDIR /app

# Install system dependencies needed for building packages
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    gcc \
    g++ \
    curl \
    git \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements.txt into the container
COPY requirements.txt .

# Upgrade pip and install Python dependencies
RUN python -m pip install --upgrade pip setuptools wheel
RUN pip install --no-cache-dir -r requirements.txt

# Copy your entire project into the container
COPY . .

# Expose port if your app uses one (Dash/Flask default)
EXPOSE 8050

# Default command to run your app (change filename if needed)
CMD ["python", "extract_data_from_screener_site.py"]
