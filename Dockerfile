# ThesisEZ Web Demo —— Render 部署用
FROM python:3.11-slim

# 装 pandoc（apt 源里 2.x 版够用）+ 中文字体（pandoc 偶尔会落字体）
RUN apt-get update && apt-get install -y --no-install-recommends \
        pandoc \
        fonts-noto-cjk \
        ca-certificates \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# 先装依赖（利用 docker layer 缓存）
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 拷代码
COPY app.py .
COPY templates/ ./templates/
COPY static/ ./static/
COPY ThesisEZ/ ./ThesisEZ/

# 清掉源码里可能残留的运行产物（保险）
RUN rm -rf ThesisEZ/__pycache__ ThesisEZ/md ThesisEZ/image \
    ThesisEZ/_ai_tags.json ThesisEZ/ThesisEZ.docx ThesisEZ/101-setup.md \
    ThesisEZ/~\$输入.docx 2>/dev/null || true

# Render 会注入 PORT 环境变量（默认 10000）
ENV PORT=8000
EXPOSE 8000

# 用 gunicorn，2 worker，超时 5 分钟（pipeline 最长 4 分钟）
CMD gunicorn --bind 0.0.0.0:$PORT --workers 2 --threads 2 --timeout 300 app:app
