# AI Presentation Builder - Multi-stage Docker Build
# Stage 1: Build stage with development dependencies
FROM python:3.11-slim as builder

# Set environment variables
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1

# Install system dependencies for building
RUN apt-get update && apt-get install -y \
    build-essential \
    gcc \
    g++ \
    && rm -rf /var/lib/apt/lists/*

# Create and set working directory
WORKDIR /app

# Copy requirements first for better layer caching
COPY requirements.txt .

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Install runtime dependencies only
RUN apt-get update && apt-get install -y \
    # Essential system packages
    curl \
    ca-certificates \
    # Image processing libraries (for python-pptx and image handling)
    libjpeg-dev \
    libpng-dev \
    # Fonts for better PowerPoint rendering
    fonts-liberation \
    fonts-dejavu-core \
    # Cleanup
    && rm -rf /var/lib/apt/lists/* \
    && apt-get clean

# Create non-root user for security
RUN groupadd -r appuser && useradd -r -g appuser appuser

# Create app directory and set ownership
WORKDIR /app
RUN chown -R appuser:appuser /app

# Copy Python packages from builder stage
COPY --from=builder /usr/local/lib/python3.11/site-packages /usr/local/lib/python3.11/site-packages
COPY --from=builder /usr/local/bin /usr/local/bin

# Copy application code
COPY --chown=appuser:appuser . .

# Create necessary directories with proper permissions
RUN mkdir -p generated_presentations assets logs && \
    chown -R appuser:appuser generated_presentations assets logs

# Switch to non-root user
USER appuser

# Expose Streamlit port
EXPOSE 8501

# Default command
CMD ["streamlit", "run", "streamlit_app.py", "--server.port=8501", "--server.address=0.0.0.0"]