# Use official Python 3.12 slim image
FROM python:3.12.2-slim

# Set working directory
WORKDIR /app

# Copy only requirements first for caching
COPY requirements.txt .

# Upgrade pip and install dependencies
RUN python -m pip install --upgrade pip setuptools wheel \
    && pip install --no-cache-dir -r requirements.txt

# Copy the rest of the project
COPY . .

# Expose the port your Dash app will run on
EXPOSE 8050

# Command to run your app
CMD ["python", "extract_data_from_screener_site.py"]
