# Деплой Telegram-бота на сервер

## 1. Подготовить сервер

На сервере нужен Docker и Docker Compose.

## 2. Загрузить проект

Скопируй на сервер файлы проекта:

- bot.py
- requirements.txt
- template.docx
- Dockerfile
- docker-compose.yml
- .env

Создай папки:

```bash
mkdir -p data contracts tmp_photos
```

Файлы `data/deals.json` и `data/users.json` бот создаст сам при первом запуске.

## 3. Настроить .env

В `.env` должны быть реальные ключи:

```text
TELEGRAM_BOT_TOKEN=...
OPENAI_API_KEY=...
OPENAI_MODEL=gpt-5.2
PAYMENTS_ENABLED=false
CLEAN_CHAT=true
```

Для ЮKassa позже:

```text
PAYMENTS_ENABLED=true
CONTRACT_PRICE_RUB=300
YOOKASSA_SHOP_ID=...
YOOKASSA_SECRET_KEY=...
YOOKASSA_RETURN_URL=https://t.me/username_бота
```

## 4. Запуск

```bash
docker compose up -d --build
```

## 5. Логи

```bash
docker logs -f telegram-contract-bot
```

## 6. Остановка

```bash
docker compose down
```
