FROM python:3.11-slim

ENV DEBIAN_FRONTEND=noninteractive

# Cài đặt LibreOffice Writer (headless) và các font chữ chuẩn tiếng Việt để chuyển đổi file .doc cũ sang .docx
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer-nogui \
    fonts-liberation \
    fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Cài đặt các thư viện Python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Sao chép toàn bộ mã nguồn ứng dụng
COPY . .

# Tạo thư mục và cấp toàn quyền ghi cho user container (Hugging Face Spaces chạy với UID 1000)
RUN mkdir -p /app/data /app/data/history /app/downloads && chmod -R 777 /app

# Cổng mặc định cho Hugging Face Spaces là 7860 (Render sẽ tự động ghi đè qua biến môi trường $PORT)
ENV PORT=7860
EXPOSE 7860

CMD ["sh", "-c", "gunicorn server:app --bind 0.0.0.0:${PORT:-7860} --workers 2 --timeout 120"]
