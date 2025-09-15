# Use official Python 3.12 slim image
FROM python:3.12.2-slim

# Set working directory inside the container
WORKDIR /app

# Copy requirements first for caching
COPY requirements.txt .

# Upgrade pip and install dependencies
RUN python -m pip install --upgrade pip setuptools wheel
RUN pip install --no-cache-dir -r requirements.txt

# Copy the rest of your project
COPY . .

# Expose the port Render will assign (default 10000)
ENV PORT 10000
EXPOSE 10000

# Set the default command to run your Dash app
CMD ["python", "wrapper.py"]
