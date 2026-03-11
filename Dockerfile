FROM python:3.12-slim

# Install nginx and supervisor
RUN apt-get update && apt-get install -y --no-install-recommends \
    nginx \
    supervisor \
    && rm -rf /var/lib/apt/lists/*

# Create non-root user
RUN useradd -m -u 1001 appuser

# Install Python dependencies
WORKDIR /app
COPY backend/requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code and frontend
COPY backend/ ./backend/
COPY frontend/ ./frontend/

# nginx config
COPY nginx/nginx.conf /etc/nginx/nginx.conf

# supervisor config
COPY supervisord.conf /etc/supervisor/conf.d/supervisord.conf

# Directories writable by appuser
RUN mkdir -p /tmp/ns_results /var/log/supervisor \
    && chown -R appuser:appuser /tmp/ns_results /var/log/supervisor \
    && chown -R appuser:appuser /var/log/nginx /var/lib/nginx \
    && mkdir -p /var/lib/nginx/body /var/lib/nginx/proxy \
    && chown -R appuser:appuser /var/lib/nginx

EXPOSE 8080

USER appuser

CMD ["/usr/bin/supervisord", "-n", "-c", "/etc/supervisor/conf.d/supervisord.conf"]
