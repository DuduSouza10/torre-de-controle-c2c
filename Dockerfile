FROM python:3.12-slim
WORKDIR /app

COPY requirements.txt /app/requirements.txt
RUN python -m pip install --no-cache-dir --upgrade pip \
    && python -m pip install --no-cache-dir -r /app/requirements.txt

COPY . /app
RUN mkdir -p /data

EXPOSE 8080
CMD ["python", "app.py"]
