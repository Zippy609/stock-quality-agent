# Use official Python 3.12 slim image 
FROM python:3.12-slim

# Set working directory
WORKDIR /app

# Copy only requirements first to leverage Docker cache
COPY requirements.txt .

# Upgrade pip, setuptools, wheel, then install dependencies
RUN python -m pip install --upgrade pip setuptools wheel \
    && pip install --no-cache-dir --root-user-action=ignore -r requirements.txt

# Copy the rest of the application
COPY . .

# Expose Render's port
ENV PORT 10000

# Command to run the Dash app
CMD ["python", "dash_wrapper.py"]
