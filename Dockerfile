FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Gunicorn: 2 workers, 60s timeout (AI parse + Sheets write typically <10s)
CMD ["gunicorn", "--bind", "0.0.0.0:5000", "--timeout", "60", "--workers", "2", "app:app"]
