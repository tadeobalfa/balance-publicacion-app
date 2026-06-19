FROM python:3.11-slim

ENV DEBIAN_FRONTEND=noninteractive
ENV PYTHONUNBUFFERED=1

RUN apt-get update && apt-get install -y \
    libreoffice \
    libreoffice-calc \
    fonts-dejavu \
    fonts-liberation \
    fontconfig \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .

RUN pip install --no-cache-dir -r requirements.txt

COPY app.py .

EXPOSE 10000

CMD streamlit run app.py --server.port=${PORT:-10000} --server.address=0.0.0.0 --server.headless=true
