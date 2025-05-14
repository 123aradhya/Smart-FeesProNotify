# Use the official Python image
FROM python:3.10-slim

# Set working directory
WORKDIR /app

# Copy project files
COPY . .

# Install dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Expose port (Render uses 10000+ dynamic port but this works fine)
EXPOSE 5000

# Start the app
CMD ["python", "app.py"]
