# 学生缺勤登记系统 — 服务器容器镜像（仅在服务器上构建运行）
FROM python:3.12-slim

WORKDIR /app

COPY requirements.txt .
# 构建时可通过 --build-arg PIP_INDEX_URL=... 指定 PyPI 镜像（国内服务器用阿里云源）
ARG PIP_INDEX_URL=https://pypi.org/simple
ENV PIP_INDEX_URL=${PIP_INDEX_URL}
RUN pip install --no-cache-dir -r requirements.txt

COPY app/ ./app/

EXPOSE 8000

CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000"]
