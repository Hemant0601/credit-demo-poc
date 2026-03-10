FROM python:3.11-slim

WORKDIR /app

# Install dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code and data
COPY demo_app.py .
COPY data/ data/

# Expose Streamlit port
EXPOSE 80

# Health check
HEALTHCHECK CMD curl --fail http://localhost:80/_stcore/health || exit 1

ENTRYPOINT ["streamlit", "run", "demo_app.py", "--server.port=80", "--server.address=0.0.0.0"]
