FROM python:3.12-slim

WORKDIR /app

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY webapp ./webapp
COPY tools ./tools

ENV HOST=0.0.0.0
ENV PORT=8787

EXPOSE 8787

CMD ["python", "webapp/server.py"]
