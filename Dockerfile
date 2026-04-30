FROM python:3.11-slim

WORKDIR /app

# Dependencias do sistema para python-docx e arquivos office
RUN apt-get update \
    && apt-get install -y --no-install-recommends gcc \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Garante pasta de saida dos documentos
RUN mkdir -p uploads

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
ENV FLASK_APP=run.py
ENV FLASK_DEBUG=0

EXPOSE 5000

CMD ["flask", "run", "--host=0.0.0.0", "--port=5000"]
