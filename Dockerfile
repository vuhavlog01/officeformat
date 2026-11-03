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

# Render tự động cấu hình cổng qua biến môi trường $PORT (mặc định 10000)
ENV PORT=10000
EXPOSE 10000

CMD ["sh", "-c", "gunicorn server:app --bind 0.0.0.0:${PORT:-10000} --workers 2 --timeout 120"]
