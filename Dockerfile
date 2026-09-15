FROM python:3.12-slim
WORKDIR /app
COPY . /app
RUN mkdir -p /data
EXPOSE 8080
CMD ["python", "app.py"]
