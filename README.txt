Какие переменные окружения нужны
--------------------------------
Все настройки сервера хранятся в файле `.env`.

Файл `.env` должен лежать в корне проекта рядом с `server.js`.

Пример содержимого `.env`:

BOT_TOKEN=123456789:YOUR_REAL_BOT_TOKEN
ADMIN_TELEGRAM_IDS=274685406
MANAGER_USERNAME=septoon
GOOGLE_SHEET_ID=1q8rEPDmvsaP6kRpzdtmbHckOpS6J6bDKSPcvRvMKO0g
GOOGLE_SERVICE_ACCOUNT_EMAIL=doncar-bot-sheets@your-project.iam.gserviceaccount.com
GOOGLE_PRIVATE_KEY="-----BEGIN PRIVATE KEY-----\nYOUR_KEY_HERE\n-----END PRIVATE KEY-----\n"
PORT=3000
BASE_URL=https://your-domain.example.com
USE_WEBHOOK=true

Подробно что означает каждая переменная
---------------------------------------

BOT_TOKEN
Токен Telegram‑бота, который выдаёт BotFather.

Пример:

BOT_TOKEN=123456789:AAHsd83kjsd83KJSD83kjsd


ADMIN_TELEGRAM_IDS
Telegram ID администратора или нескольких администраторов.

Если админ один:

ADMIN_TELEGRAM_IDS=274685406

Если несколько:

ADMIN_TELEGRAM_IDS=274685406,123456789


MANAGER_USERNAME
Username менеджера Telegram без символа @.

Пример:

MANAGER_USERNAME=septoon


GOOGLE_SHEET_ID
ID Google таблицы.

Его можно взять из URL таблицы.

Например URL:

https://docs.google.com/spreadsheets/d/1q8rEPDmvsaP6kRpzdtmbHckOpS6J6bDKSPcvRvMKO0g/edit

ID таблицы будет:

1q8rEPDmvsaP6kRpzdtmbHckOpS6J6bDKSPcvRvMKO0g


GOOGLE_SERVICE_ACCOUNT_EMAIL
Email сервисного аккаунта из JSON файла Google Cloud.

Пример из JSON:

"client_email": "doncar-bot@project-id.iam.gserviceaccount.com"

Тогда в `.env` нужно написать:

GOOGLE_SERVICE_ACCOUNT_EMAIL=doncar-bot@project-id.iam.gserviceaccount.com


GOOGLE_PRIVATE_KEY
Это приватный ключ из JSON файла сервисного аккаунта.

В JSON он выглядит так:

"private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQD...\n...\n-----END PRIVATE KEY-----\n"

Очень важно:

В `.env` ключ должен быть записан **в одну строку** и переносы строк должны быть заменены на `\\n`.

Правильный формат в `.env`:

GOOGLE_PRIVATE_KEY="-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQD...\n...\n-----END PRIVATE KEY-----\n"

Типичные ошибки:

❌ Нельзя вставлять ключ в несколько строк

Неправильно:

GOOGLE_PRIVATE_KEY="-----BEGIN PRIVATE KEY-----
MIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQD
...
-----END PRIVATE KEY-----"


❌ Нельзя удалять `BEGIN PRIVATE KEY` или `END PRIVATE KEY`

❌ Нельзя убирать `\\n`


PORT
Порт на котором будет работать сервер.

Обычно:

PORT=3000


BASE_URL
Публичный URL сервера.

Нужен только если используется webhook.

Пример:

BASE_URL=https://doncar-bot.onrender.com


USE_WEBHOOK
Режим работы бота.

true — бот работает через webhook (для VPS / Render / Railway)

false — бот работает через polling (для локальной разработки)

Пример для локальной разработки:

USE_WEBHOOK=false

Пример для продакшена:

USE_WEBHOOK=true


Итог: какие данные нужно вставить из Google Cloud JSON
------------------------------------------------------

Из скачанного JSON сервисного аккаунта нужны только два поля:

client_email
private_key

client_email → GOOGLE_SERVICE_ACCOUNT_EMAIL

private_key → GOOGLE_PRIVATE_KEY

Остальные поля из JSON не используются.