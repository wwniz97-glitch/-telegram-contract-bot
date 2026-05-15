FROM python:3.12-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY bot.py template.docx ./

RUN mkdir -p contracts tmp_photos data

CMD ["python", "bot.py"]