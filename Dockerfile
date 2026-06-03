# Maquina Portfolio Dashboard — slim Python image.
# Unlike the Ember app, this service has no WeasyPrint/Pango dependency,
# so a plain slim base is sufficient.
FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Railway provides $PORT at runtime; gunicorn.conf.py reads it.
EXPOSE 8080
CMD ["gunicorn", "-c", "gunicorn.conf.py", "app:app"]
