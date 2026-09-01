FROM python:3.11-slim

WORKDIR /app

# -------------------------------
# Install system dependencies
# -------------------------------
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    gcc \
    libmupdf-dev \
    && rm -rf /var/lib/apt/lists/*

# -------------------------------
# Python settings
# -------------------------------
ENV PYTHONUNBUFFERED=1

# -------------------------------
# Install Python dependencies
# -------------------------------
COPY requirements.txt .

RUN pip install --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt

# -------------------------------
# Copy application code
# -------------------------------
COPY . .

# -------------------------------
# Expose FastAPI port
# -------------------------------
EXPOSE 8000

# -------------------------------
# Run FastAPI app
# -------------------------------
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]