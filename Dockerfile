# Estágio 1: Configuração do Kali Linux e Google Chrome
FROM debian:buster-slim as kali

RUN apt-get update && \
    apt-get upgrade -y && \
    apt-get -y install wget gnupg xorg xauth libnss3 libgconf-2-4 libfontconfig1

RUN wget -q -O - https://dl-ssl.google.com/linux/linux_signing_key.pub | apt-key add - && \
    echo "deb [arch=amd64] http://dl.google.com/linux/chrome/deb/ stable main" >> /etc/apt/sources.list.d/google-chrome.list && \
    apt-get update && \
    apt-get install -y google-chrome-stable

ENV DBUS_SESSION_BUS_ADDRESS=/dev/null

# Estágio 2: Configuração do ambiente Python
FROM python:3.9

WORKDIR /app

COPY requirements.txt /app/

RUN pip install --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt && \
    pip install selenium webdriver-manager && \
    pip install telegram && \
    pip install python-telegram-bot && \
    pip install pandas && \
    pip install openpyxl

COPY . /app/

# Comando para iniciar o script monitoramento.py
CMD ["python", "app-telegram-monitoramento.py"]
