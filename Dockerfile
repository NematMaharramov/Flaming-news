FROM python:3.11-slim

# WeasyPrint (used for PDF export) needs Pango/Cairo/GDK-Pixbuf at the
# system level — these aren't part of the plain Python buildpack, so we
# build via Docker to guarantee they're present in production.
RUN apt-get update && apt-get install -y --no-install-recommends \
    libpango-1.0-0 \
    libpangocairo-1.0-0 \
    libpangoft2-1.0-0 \
    libcairo2 \
    libgdk-pixbuf-2.0-0 \
    libffi-dev \
    shared-mime-info \
    fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .

# Daily JSON + photo storage — persisted via a Render disk mounted here (see render.yaml)
RUN mkdir -p /app/data

EXPOSE 10000
CMD ["gunicorn", "app:app", "--bind", "0.0.0.0:10000"]
