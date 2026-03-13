import 'dotenv/config';
import express from 'express';
import TelegramBot from 'node-telegram-bot-api';
import { google } from 'googleapis';

const APP_VERSION = 'v1.0.0';

// ===== ENV =====
const {
  BOT_TOKEN,
  ADMIN_TELEGRAM_IDS = '',
  MANAGER_USERNAME = 'septoon',
  GOOGLE_SHEET_ID,
  GOOGLE_SERVICE_ACCOUNT_EMAIL,
  GOOGLE_PRIVATE_KEY,
  PORT = '3000',
  BASE_URL = '',
  USE_WEBHOOK = 'true',
} = process.env;

if (!BOT_TOKEN) throw new Error('BOT_TOKEN is required');
if (!GOOGLE_SHEET_ID) throw new Error('GOOGLE_SHEET_ID is required');
if (!GOOGLE_SERVICE_ACCOUNT_EMAIL) throw new Error('GOOGLE_SERVICE_ACCOUNT_EMAIL is required');
if (!GOOGLE_PRIVATE_KEY) throw new Error('GOOGLE_PRIVATE_KEY is required');

const adminIds = ADMIN_TELEGRAM_IDS.split(',')
  .map(v => v.trim())
  .filter(Boolean);

const SHEET_CLIENTS = 'clients';
const SHEET_HISTORY = 'history';
const SHEET_STATES = 'states';
const SHEET_REQUESTS = 'requests';

const STATES = {
  NONE: '',
  WAITING_NAME: 'waiting_name',
  WAITING_PHONE_CONTACT: 'waiting_phone_contact',
  CLIENT_WAITING_RENTAL_AMOUNT_FOR_REDEEM: 'client_waiting_rental_amount_for_redeem',
  CLIENT_WAITING_REDEEM_DATETIME: 'client_waiting_redeem_datetime',
  CLIENT_WAITING_REDEEM_RENTAL_ID: 'client_waiting_redeem_rental_id',
  ADMIN_CLIENT_ACTIONS: 'admin_client_actions',
  ADMIN_WAITING_CLIENT_IDENTIFIER: 'admin_waiting_client_identifier',
  ADMIN_WAITING_ACCRUAL_AMOUNT: 'admin_waiting_accrual_amount',
  ADMIN_WAITING_ACCRUAL_DATETIME: 'admin_waiting_accrual_datetime',
  ADMIN_WAITING_ACCRUAL_RENTAL_ID: 'admin_waiting_accrual_rental_id',
  ADMIN_WAITING_HISTORY_IDENTIFIER: 'admin_waiting_history_identifier',
  ADMIN_WAITING_FIND_IDENTIFIER: 'admin_waiting_find_identifier',
  ADMIN_WAITING_BALANCE_IDENTIFIER: 'admin_waiting_balance_identifier',
  ADMIN_WAITING_MANUAL_ACCRUAL_IDENTIFIER: 'admin_waiting_manual_accrual_identifier',
  ADMIN_WAITING_MANUAL_ACCRUAL_AMOUNT: 'admin_waiting_manual_accrual_amount',
  ADMIN_WAITING_MANUAL_ACCRUAL_COMMENT: 'admin_waiting_manual_accrual_comment',
  ADMIN_WAITING_MANUAL_REDEEM_IDENTIFIER: 'admin_waiting_manual_redeem_identifier',
  ADMIN_WAITING_MANUAL_REDEEM_AMOUNT: 'admin_waiting_manual_redeem_amount',
  ADMIN_WAITING_MANUAL_REDEEM_COMMENT: 'admin_waiting_manual_redeem_comment',
  ADMIN_WAITING_MANUAL_CLIENT_IDENTIFIER: 'admin_waiting_manual_client_identifier',
  ADMIN_WAITING_MANUAL_AMOUNT: 'admin_waiting_manual_amount',
  ADMIN_WAITING_MANUAL_COMMENT: 'admin_waiting_manual_comment',
};

const REQUEST_STATUS = {
  PENDING: 'pending',
  APPROVED: 'approved',
  REJECTED: 'rejected',
  DUPLICATE: 'duplicate',
};

const OPERATION_TYPE = {
  ACCRUAL: 'accrual',
  REDEEM: 'redeem',
  MANUAL_ACCRUAL: 'manual_accrual',
  MANUAL_REDEEM: 'manual_redeem',
};

const BUTTONS = {
  MY_CARD: 'Моя карта',
  MY_BALANCE: 'Мой баланс',
  USE_BONUSES: 'Использовать бонусы',
  CONTACT_MANAGER: 'Связаться с менеджером',
  ADMIN_MENU: 'Админ-меню',
  FIND_CLIENT: 'Найти клиента',
  CLIENT_BALANCE: 'Баланс клиента',
  CLIENT_HISTORY: 'История клиента',
  ACCRUE_BONUSES: 'Начислить бонусы',
  MANUAL_ACCRUAL: 'Начислить вручную',
  MANUAL_REDEEM: 'Списать вручную',
  REDEEM_REQUESTS: 'Заявки на списание',
};

const CALLBACKS = {
  APPROVE_REQUEST: 'approve_request:',
  REJECT_REQUEST: 'reject_request:',
};

const CACHE_LIMIT = 1000;
const SHEET_CACHE_TTLS = {
  [SHEET_CLIENTS]: 10_000,
  [SHEET_HISTORY]: 5_000,
  [SHEET_REQUESTS]: 5_000,
  [SHEET_STATES]: 3_000,
};

const sheetCache = {
  [SHEET_CLIENTS]: { data: null, expiresAt: 0 },
  [SHEET_HISTORY]: { data: null, expiresAt: 0 },
  [SHEET_REQUESTS]: { data: null, expiresAt: 0 },
  [SHEET_STATES]: { data: null, expiresAt: 0 },
};

const clientIndexesCache = new WeakMap();

// ===== GOOGLE SHEETS =====
const auth = new google.auth.JWT({
  email: GOOGLE_SERVICE_ACCOUNT_EMAIL,
  key: GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

const sheets = google.sheets({ version: 'v4', auth });

function getSheetCacheEntry(sheetName) {
  return sheetCache[sheetName] || null;
}

function getSheetCacheTtl(sheetName) {
  return SHEET_CACHE_TTLS[sheetName] || 0;
}

function extractSheetNameFromRange(range) {
  const rawRange = safeString(range).trim();
  const sheetPart = rawRange.split('!')[0] || '';
  return sheetPart.replace(/^'/, '').replace(/'$/, '');
}

function setCachedSheetRows(sheetName, rows, ttlMs) {
  const cacheEntry = getSheetCacheEntry(sheetName);
  if (!cacheEntry || !ttlMs) return rows;

  cacheEntry.data = rows;
  cacheEntry.expiresAt = Date.now() + ttlMs;
  return rows;
}

function invalidateSheetCache(sheetName) {
  const cacheEntry = getSheetCacheEntry(sheetName);
  if (!cacheEntry) return;

  cacheEntry.data = null;
  cacheEntry.expiresAt = 0;
  logEvent('CACHE INVALIDATE', { sheet: sheetName });
}

async function getCachedSheetRows(sheetName) {
  const cacheEntry = getSheetCacheEntry(sheetName);
  const ttlMs = getSheetCacheTtl(sheetName);

  if (cacheEntry?.data && cacheEntry.expiresAt > Date.now()) {
    logEvent('CACHE HIT', { sheet: sheetName });
    return cacheEntry.data;
  }

  logEvent('CACHE MISS', { sheet: sheetName });
  const rows = await readRange(`${sheetName}!A:Z`);
  return setCachedSheetRows(sheetName, rows, ttlMs);
}

async function ensureSheet(title, headers) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId: GOOGLE_SHEET_ID });
  const exists = meta.data.sheets?.some(s => s.properties?.title === title);

  if (!exists) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: GOOGLE_SHEET_ID,
      requestBody: {
        requests: [{ addSheet: { properties: { title } } }],
      },
    });
  }

  const values = await readRange(`${title}!A1:Z1`);
  const row = values[0] || [];
  const isEmpty = row.length === 0 || row.every(v => safeString(v).trim() === '');

  if (isEmpty) {
    await writeRange(`${title}!A1`, [headers]);
  }
}

async function initSheets() {
  await ensureSheet(SHEET_CLIENTS, [
    'telegram_id',
    'name',
    'phone',
    'bonus_balance',
    'created_at',
    'updated_at',
  ]);

  await ensureSheet(SHEET_HISTORY, [
    'operation_id',
    'date',
    'telegram_id',
    'type',
    'amount',
    'comment',
    'admin_id',
    'rental_id',
    'rental_datetime',
    'duplicate_key',
  ]);

  await ensureSheet(SHEET_STATES, [
    'telegram_id',
    'state',
    'temp_value',
    'updated_at',
  ]);

  await ensureSheet(SHEET_REQUESTS, [
    'request_id',
    'created_at',
    'telegram_id',
    'phone',
    'rental_amount',
    'max_allowed_bonus',
    'requested_bonus',
    'status',
    'admin_id',
    'rental_id',
    'rental_datetime',
  ]);
}

async function readRange(range) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: GOOGLE_SHEET_ID,
    range,
  });
  return res.data.values || [];
}

async function writeRange(range, values) {
  await sheets.spreadsheets.values.update({
    spreadsheetId: GOOGLE_SHEET_ID,
    range,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values },
  });

  const sheetName = extractSheetNameFromRange(range);
  if (sheetName) invalidateSheetCache(sheetName);
}

async function appendRow(sheetName, row) {
  await sheets.spreadsheets.values.append({
    spreadsheetId: GOOGLE_SHEET_ID,
    range: `${sheetName}!A:Z`,
    valueInputOption: 'USER_ENTERED',
    insertDataOption: 'INSERT_ROWS',
    requestBody: { values: [row] },
  });

  invalidateSheetCache(sheetName);
}

async function getAllRows(sheetName) {
  return await getCachedSheetRows(sheetName);
}

async function updateRowCells(sheetName, rowNumber, startColLetter, values) {
  await writeRange(`${sheetName}!${startColLetter}${rowNumber}`, [values]);
}

// ===== HELPERS =====
function nowIso() {
  return new Date().toISOString();
}

function safeString(value, fallback = '') {
  if (value === undefined || value === null) return fallback;
  return String(value);
}

function safeNumber(value, fallback = 0) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function safeJsonParse(value, fallback = {}) {
  try {
    if (!value) return fallback;
    return JSON.parse(value);
  } catch {
    return fallback;
  }
}

function round2(n) {
  return Math.round(safeNumber(n) * 100) / 100;
}

function pad2(n) {
  return safeString(n).padStart(2, '0');
}

function formatLogValue(value) {
  if (typeof value === 'string') return value;
  return JSON.stringify(value);
}

function logEvent(tag, payload = {}) {
  const details = Object.entries(payload)
    .filter(([, value]) => value !== undefined && value !== null && value !== '')
    .map(([key, value]) => `${key}=${formatLogValue(value)}`)
    .join(' ');

  console.log(`[${tag}]${details ? ` ${details}` : ''}`);
}

function logError(tag, error, payload = {}) {
  const details = Object.entries(payload)
    .filter(([, value]) => value !== undefined && value !== null && value !== '')
    .map(([key, value]) => `${key}=${formatLogValue(value)}`)
    .join(' ');

  console.error(`[${tag}]${details ? ` ${details}` : ''}`, error);
}

function rememberProcessed(set, key, limit = CACHE_LIMIT) {
  const normalizedKey = safeString(key);
  if (!normalizedKey) return;

  set.add(normalizedKey);
  while (set.size > limit) {
    const oldest = set.values().next().value;
    if (!oldest) break;
    set.delete(oldest);
  }
}

function parseMoney(value) {
  const n = safeNumber(safeString(value).replace(',', '.').replace(/\s/g, ''), NaN);
  return Number.isFinite(n) ? round2(n) : null;
}

function normalizePhone(phone) {
  const cleaned = safeString(phone).replace(/[^\d+]/g, '');
  if (/^\+7\d{10}$/.test(cleaned)) return cleaned;
  if (/^8\d{10}$/.test(cleaned)) return '+7' + cleaned.slice(1);
  if (/^7\d{10}$/.test(cleaned)) return '+' + cleaned;
  return null;
}

function formatPhoneDisplay(phone) {
  const rawPhone = safeString(phone).trim();
  if (!rawPhone) return '-';

  const digits = rawPhone.replace(/\D/g, '');
  if (/^7\d{10}$/.test(digits)) return `+${digits}`;
  if (/^8\d{10}$/.test(digits)) return digits;

  return rawPhone;
}

function normalizeDateTime(value) {
  const str = safeString(value).trim();
  if (!str) return null;

  const dotMatch = str.match(/^(\d{2})\.(\d{2})\.(\d{4})(?:\s+(\d{2}):(\d{2}))?$/);
  const isoMatch = str.match(/^(\d{4})-(\d{2})-(\d{2})(?:\s+(\d{2}):(\d{2}))?$/);

  let day;
  let month;
  let year;
  let hours = '00';
  let minutes = '00';

  if (dotMatch) {
    [, day, month, year, hours = '00', minutes = '00'] = dotMatch;
  } else if (isoMatch) {
    [, year, month, day, hours = '00', minutes = '00'] = isoMatch;
  } else {
    return null;
  }

  const parsed = new Date(
    safeNumber(year),
    safeNumber(month) - 1,
    safeNumber(day),
    safeNumber(hours),
    safeNumber(minutes)
  );

  if (
    Number.isNaN(parsed.getTime()) ||
    parsed.getFullYear() !== safeNumber(year) ||
    parsed.getMonth() !== safeNumber(month) - 1 ||
    parsed.getDate() !== safeNumber(day) ||
    parsed.getHours() !== safeNumber(hours) ||
    parsed.getMinutes() !== safeNumber(minutes)
  ) {
    return null;
  }

  return `${pad2(day)}.${pad2(month)}.${year} ${pad2(hours)}:${pad2(minutes)}`;
}

function formatDatePrompt() {
  return 'Пример: 13.03.2026 18:30';
}

const TEXT = {
  ACCESS_DENIED_ADMIN: 'У вас нет доступа к админ-функциям.',
  ACCESS_DENIED_ACTION: 'У вас нет доступа к этому действию.',
  ACTION_PROMPT: 'Выберите действие',
  ADMIN_MENU: 'Админ-меню',
  CALLBACK_ACTION_ERROR: 'Произошла ошибка при обработке действия.',
  CLIENT_FOUND: 'Клиент найден',
  CLIENT_NOT_FOUND: 'Клиент не найден.\nИспользуйте Telegram ID или телефон.',
  CLIENT_NOT_FOUND_RESTART: 'Клиент не найден. Начните заново.',
  CLIENT_NOT_FOUND_REGISTER: 'Клиент не найден. Пройдите /start заново.',
  COMMENT_REQUIRED: 'Комментарий не должен быть пустым.',
  CONTACT_NOT_REQUESTED: 'Сейчас номер телефона не запрашивался.',
  DATE_PROMPT: `Введите дату и время аренды\n${formatDatePrompt()}`,
  DUPLICATE_ACCRUAL: 'Похоже, это дубль начисления. Операция отменена.',
  DUPLICATE_MANUAL_OPERATION: 'Похоже, такая операция уже есть. Проверьте сумму и комментарий.',
  DUPLICATE_REDEEM: 'Похоже, это дубль списания. Подтверждение остановлено.',
  INVALID_AMOUNT: 'Введите корректную сумму.',
  INVALID_NAME: 'Введите корректное имя.',
  INVALID_PHONE: 'Не удалось распознать номер телефона.',
  INVALID_RENTAL_AMOUNT: 'Введите корректную сумму аренды.',
  NO_HISTORY: 'Операций пока нет.',
  NO_PENDING_REQUESTS: 'Нет заявок на списание.',
  NOT_ENOUGH_BONUSES: 'Недостаточно бонусов для списания.\nВведите другую сумму.',
  REGISTER_COMPLETE: 'Регистрация завершена.',
  REGISTER_FIRST: 'Сначала зарегистрируйтесь через /start',
  REQUEST_ALREADY_PROCESSED_PREFIX: 'Заявка уже обработана. Статус:',
  REQUEST_NOT_FOUND: 'Заявка не найдена.',
  REQUEST_REJECTED_CLIENT: 'Ваш запрос на списание бонусов был отклонён менеджером.',
  REQUEST_RENTAL_ID_PROMPT: 'Теперь отправьте номер аренды/брони.\nЕсли его нет — отправьте "-".',
  REQUEST_RENTAL_ID_PROMPT_ADMIN: 'Введите номер аренды/брони. Если его нет — отправьте "-".',
  SAME_REQUEST_EXISTS: 'Похожая заявка уже существует. Проверьте, не отправляли ли вы её ранее.',
  SEND_CONTACT_PROMPT: 'Теперь нажмите кнопку ниже, чтобы отправить ваш номер телефона.',
  SEND_OWN_CONTACT: 'Нужно отправить именно свой контакт.',
  SERVICE_UNAVAILABLE: 'Сервис временно недоступен. Попробуйте позже.',
  UNKNOWN_COMMAND: 'Не понял команду. Нажмите /start',
  WAIT_COMMENT: 'Введите комментарий.',
  WAIT_IDENTIFIER: 'Введите Telegram ID клиента или его телефон',
  WAIT_MANUAL_ACCRUAL_AMOUNT: 'Введите сумму бонусов.',
  WAIT_MANUAL_REDEEM_AMOUNT: 'Введите сумму списания.',
  WAIT_RENTAL_AMOUNT: 'Введите сумму аренды.',
  ZERO_REDEEM: 'Списывать нечего: баланс пустой или лимит исчерпан.',
};

function formatDateTimeValue(value) {
  const normalized = normalizeDateTime(value);
  if (normalized) return normalized;

  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return safeString(value);

  return `${pad2(date.getDate())}.${pad2(date.getMonth() + 1)}.${date.getFullYear()} ${pad2(date.getHours())}:${pad2(date.getMinutes())}`;
}

function formatHistoryLine(item) {
  const typeLabel =
    item.type === OPERATION_TYPE.ACCRUAL
      ? 'Начисление'
      : item.type === OPERATION_TYPE.REDEEM
      ? 'Списание'
      : item.type === OPERATION_TYPE.MANUAL_ACCRUAL
      ? 'Ручное начисление'
      : item.type === OPERATION_TYPE.MANUAL_REDEEM
      ? 'Ручное списание'
      : item.type;

  const comment = item.comment ? ` — ${item.comment}` : '';
  const rental = item.rental_id ? ` — аренда: ${item.rental_id}` : '';
  const date = formatDateTimeValue(item.rental_datetime || item.date || '');

  return `${date} • ${typeLabel} • ${item.amount} бонусов${rental}${comment}`;
}

function formatClientCard(client) {
  return (
    `Имя: ${client.name || '-'}\n` +
    `Телефон: ${formatPhoneDisplay(client.phone)}\n` +
    `Telegram ID: ${client.telegram_id}\n` +
    `Баланс: ${round2(client.bonus_balance)} бонусов`
  );
}

function formatClientHistory(history) {
  return history.length ? history.map(formatHistoryLine).join('\n') : TEXT.NO_HISTORY;
}

function formatRequestCard(request, client) {
  return [
    `request_id: ${request.request_id}`,
    `Имя клиента: ${client?.name || '-'}`,
    `Телефон: ${formatPhoneDisplay(request.phone || client?.phone)}`,
    `Telegram ID: ${request.telegram_id}`,
    `Аренда: ${request.rental_amount} ₽`,
    `Дата/время аренды: ${request.rental_datetime}`,
    `Номер аренды: ${request.rental_id || '-'}`,
    `Доступно к списанию: ${request.requested_bonus}`,
    `Статус: ${request.status || REQUEST_STATUS.PENDING}`,
  ].join('\n');
}

function formatBalanceMessage(client) {
  return formatClientCard(client);
}

function formatAccrualMessage(bonus, balance) {
  return `Начислено ${bonus} бонусов\nНовый баланс клиента: ${balance}`;
}

function formatRedeemMessage(amount, balance) {
  return `Списано ${amount} бонусов\nНовый баланс клиента: ${balance}`;
}

function formatClientSummaryMessage(client, history) {
  return `${TEXT.CLIENT_FOUND}\n\n${formatClientCard(client)}\n\nИстория операций:\n\n${formatClientHistory(history)}`;
}

function formatFoundClientMessage(client) {
  return `${TEXT.CLIENT_FOUND}\n\n${formatClientCard(client)}`;
}

function formatClientPromptMessage(client, prompt) {
  return `${formatFoundClientMessage(client)}\n\n${prompt}`;
}

function formatRequestProcessedMessage(status) {
  return `${TEXT.REQUEST_ALREADY_PROCESSED_PREFIX} ${status}`;
}

function formatAdminRentalCalculation(amount, client, maxAllowed, requested) {
  return (
    `Сумма аренды: ${amount} ₽\n` +
    `Ваш баланс: ${round2(client.bonus_balance)} бонусов\n` +
    `Максимум по правилу 10%: ${maxAllowed}\n` +
    `Доступно к списанию: ${requested}\n\n` +
    TEXT.DATE_PROMPT
  );
}

function formatNoRedeemAvailableMessage(client) {
  return `Ваш баланс: ${round2(client.bonus_balance)} бонусов.\nДля этой аренды списание недоступно.`;
}

function formatRequestCreatedMessage(requestId, requestedBonus) {
  return (
    `Заявка создана.\n` +
    `request_id: ${requestId}\n` +
    `К списанию запрошено: ${requestedBonus} бонусов\n` +
    `Ожидайте подтверждения менеджера.`
  );
}

function formatRequestRejectedMessage(requestId) {
  return `Заявка ${requestId} отклонена.`;
}

function buildRequestId(telegramId) {
  return `REQ-${telegramId}-${Date.now()}`;
}

function buildOperationId(type) {
  return `OP-${type}-${Date.now()}-${Math.floor(Math.random() * 1000)}`;
}

function buildDuplicateKey(type, telegramId, rentalAmount, rentalDateTime, rentalId) {
  if (rentalId) return [type, telegramId, safeString(rentalId)].join('|');
  return [type, telegramId, safeString(rentalAmount), safeString(rentalDateTime)].join('|');
}

function buildManualOperationDuplicateKey(type, telegramId, amount, comment) {
  return [type, telegramId, safeString(amount), safeString(comment).trim().toLowerCase()].join('|');
}

function buildManualAccrualDuplicateKey(telegramId, amount, comment) {
  return buildManualOperationDuplicateKey('manual_accrual', telegramId, amount, comment);
}

function buildManualRedeemDuplicateKey(telegramId, amount, comment) {
  return buildManualOperationDuplicateKey('manual_redeem', telegramId, amount, comment);
}

function currentDisplayDateTime() {
  return formatDateTimeValue(nowIso());
}

function normalizeDuplicateKeyValue(duplicateKey) {
  const key = safeString(duplicateKey);
  const parts = key.split('|');

  if (parts.length === 4) {
    const normalizedDateTime = normalizeDateTime(parts[3]);
    if (normalizedDateTime) {
      return [parts[0], parts[1], parts[2], normalizedDateTime].join('|');
    }
  }

  return key;
}

function dateTimeToTimestamp(value) {
  const normalized = normalizeDateTime(value);
  if (normalized) {
    const match = normalized.match(/^(\d{2})\.(\d{2})\.(\d{4}) (\d{2}):(\d{2})$/);
    if (match) {
      const [, day, month, year, hours, minutes] = match;
      return new Date(
        safeNumber(year),
        safeNumber(month) - 1,
        safeNumber(day),
        safeNumber(hours),
        safeNumber(minutes)
      ).getTime();
    }
  }

  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? 0 : date.getTime();
}

function isAdmin(telegramId) {
  return adminIds.includes(safeString(telegramId));
}

function mapClientRow(row, rowNumber) {
  return {
    rowNumber,
    telegram_id: safeString(row[0]),
    name: row[1] || '',
    phone: row[2] || '',
    bonus_balance: safeNumber(row[3]),
    created_at: row[4] || '',
    updated_at: row[5] || '',
  };
}

function buildClientIndexes(rows) {
  const byTelegramId = new Map();
  const byPhone = new Map();

  for (let i = 1; i < rows.length; i++) {
    const client = mapClientRow(rows[i], i + 1);
    if (client.telegram_id) byTelegramId.set(client.telegram_id, client);

    const normalizedPhone = normalizePhone(client.phone) || safeString(client.phone);
    if (normalizedPhone) byPhone.set(normalizedPhone, client);
  }

  return { byTelegramId, byPhone };
}

function getClientIndexes(rows) {
  const cachedIndexes = clientIndexesCache.get(rows);
  if (cachedIndexes) return cachedIndexes;

  const indexes = buildClientIndexes(rows);
  clientIndexesCache.set(rows, indexes);
  return indexes;
}

function mapRequestRow(row, rowNumber) {
  return {
    rowNumber,
    request_id: row[0],
    created_at: row[1] || '',
    telegram_id: safeString(row[2]),
    phone: row[3] || '',
    rental_amount: safeNumber(row[4]),
    max_allowed_bonus: safeNumber(row[5]),
    requested_bonus: safeNumber(row[6]),
    status: row[7] || '',
    admin_id: row[8] || '',
    rental_id: row[9] || '',
    rental_datetime: row[10] || '',
  };
}

function mapHistoryRow(row, index) {
  return {
    _sort_index: index,
    operation_id: row[0] || '',
    date: row[1] || '',
    telegram_id: safeString(row[2]),
    type: row[3] || '',
    amount: safeNumber(row[4]),
    comment: row[5] || '',
    admin_id: row[6] || '',
    rental_id: row[7] || '',
    rental_datetime: row[8] || '',
    duplicate_key: row[9] || '',
  };
}

function serializeStateValue(tempValue = {}) {
  if (typeof tempValue === 'string') return tempValue;
  return JSON.stringify(tempValue || {});
}

// ===== DATA ACCESS =====
async function getClientByTelegramId(telegramId) {
  const rows = await getAllRows(SHEET_CLIENTS);
  const indexes = getClientIndexes(rows);
  return indexes.byTelegramId.get(safeString(telegramId)) || null;
}

async function getClientByPhone(phone) {
  const normalized = normalizePhone(phone);
  if (!normalized) return null;

  const rows = await getAllRows(SHEET_CLIENTS);
  const indexes = getClientIndexes(rows);
  return indexes.byPhone.get(normalized) || null;
}

async function findClientByIdentifier(identifier) {
  const byId = await getClientByTelegramId(identifier);
  if (byId) return byId;
  return await getClientByPhone(identifier);
}

async function upsertClient(telegramId, name, phone) {
  const normalizedPhone = normalizePhone(phone);
  const existing = await getClientByTelegramId(telegramId);
  const now = nowIso();

  if (existing) {
    await updateRowCells(SHEET_CLIENTS, existing.rowNumber, 'B', [
      name,
      normalizedPhone,
      existing.bonus_balance,
      existing.created_at || now,
      now,
    ]);
    return;
  }

  await appendRow(SHEET_CLIENTS, [
    telegramId,
    name,
    normalizedPhone,
    0,
    now,
    now,
  ]);
}

async function changeBonusBalance(telegramId, delta) {
  const client = await getClientByTelegramId(telegramId);
  if (!client) throw new Error(`Client not found: ${telegramId}`);

  const next = round2(Math.max(0, client.bonus_balance + delta));
  await updateRowCells(SHEET_CLIENTS, client.rowNumber, 'D', [next, nowIso()]);
  return next;
}

async function addHistory(row) {
  await appendRow(SHEET_HISTORY, [
    row.operation_id,
    row.date || nowIso(),
    row.telegram_id,
    row.type,
    row.amount,
    row.comment,
    row.admin_id || '',
    row.rental_id || '',
    row.rental_datetime || '',
    row.duplicate_key || '',
  ]);
}

async function historyDuplicateKeyExists(duplicateKey) {
  const normalizedTarget = normalizeDuplicateKeyValue(duplicateKey);
  const rows = await getAllRows(SHEET_HISTORY);
  for (let i = 1; i < rows.length; i++) {
    if (normalizeDuplicateKeyValue(rows[i][9]) === normalizedTarget) return true;
  }
  return false;
}

async function addRequest(request) {
  await appendRow(SHEET_REQUESTS, [
    request.request_id,
    nowIso(),
    request.telegram_id,
    request.phone,
    request.rental_amount,
    request.max_allowed_bonus,
    request.requested_bonus,
    request.status,
    request.admin_id || '',
    request.rental_id || '',
    request.rental_datetime || '',
  ]);
}

async function getRequestById(requestId) {
  const rows = await getAllRows(SHEET_REQUESTS);
  for (let i = 1; i < rows.length; i++) {
    if (safeString(rows[i][0]) === safeString(requestId)) {
      return mapRequestRow(rows[i], i + 1);
    }
  }
  return null;
}

async function getClientHistory(telegramId, limit = 10) {
  const rows = await getAllRows(SHEET_HISTORY);
  const history = [];

  for (let i = 1; i < rows.length; i++) {
    if (safeString(rows[i][2]) !== safeString(telegramId)) continue;
    history.push(mapHistoryRow(rows[i], i));
  }

  return history
    .sort((a, b) => {
      const byDate = dateTimeToTimestamp(b.date) - dateTimeToTimestamp(a.date);
      if (byDate !== 0) return byDate;
      return b._sort_index - a._sort_index;
    })
    .slice(0, limit)
    .map(({ _sort_index, ...item }) => item);
}

async function getPendingRequests() {
  const rows = await getAllRows(SHEET_REQUESTS);
  const result = [];
  for (let i = 1; i < rows.length; i++) {
    if (safeString(rows[i][7]) === REQUEST_STATUS.PENDING) {
      result.push(mapRequestRow(rows[i], i + 1));
    }
  }
  return result;
}

async function markRequestStatus(requestId, status, adminId) {
  const req = await getRequestById(requestId);
  if (!req) return;
  await updateRowCells(SHEET_REQUESTS, req.rowNumber, 'H', [status, adminId || '']);
}

async function requestLooksDuplicate(telegramId, rentalAmount, rentalDateTime, rentalId) {
  const rows = await getAllRows(SHEET_REQUESTS);
  for (let i = 1; i < rows.length; i++) {
    const sameClient = safeString(rows[i][2]) === safeString(telegramId);
    const sameAmount = safeNumber(rows[i][4]) === safeNumber(rentalAmount);
    const sameDateTime = normalizeDateTime(rows[i][10]) === normalizeDateTime(rentalDateTime);
    const sameRentalId = safeString(rows[i][9]) === safeString(rentalId);
    const activeStatus = [REQUEST_STATUS.PENDING, REQUEST_STATUS.APPROVED].includes(safeString(rows[i][7]));

    if (!activeStatus) continue;
    if (rentalId && sameClient && sameRentalId) return true;
    if (!rentalId && sameClient && sameAmount && sameDateTime) return true;
  }
  return false;
}

async function getState(telegramId) {
  const rows = await getAllRows(SHEET_STATES);
  for (let i = 1; i < rows.length; i++) {
    if (safeString(rows[i][0]) === safeString(telegramId)) {
      return rows[i][1] || '';
    }
  }
  return STATES.NONE;
}

async function setState(telegramId, state, tempValue) {
  const rows = await getAllRows(SHEET_STATES);
  const now = nowIso();
  const serialized = serializeStateValue(tempValue);

  for (let i = 1; i < rows.length; i++) {
    if (safeString(rows[i][0]) === safeString(telegramId)) {
      await updateRowCells(SHEET_STATES, i + 1, 'B', [state, serialized, now]);
      logEvent('STATE', { telegramId, state: state || 'none' });
      return;
    }
  }

  await appendRow(SHEET_STATES, [telegramId, state, serialized, now]);
  logEvent('STATE', { telegramId, state: state || 'none' });
}

async function parseStateData(telegramId) {
  const rows = await getAllRows(SHEET_STATES);
  for (let i = 1; i < rows.length; i++) {
    if (safeString(rows[i][0]) === safeString(telegramId)) {
      return safeJsonParse(rows[i][2], {});
    }
  }
  return {};
}

async function clearState(telegramId) {
  const rows = await getAllRows(SHEET_STATES);
  for (let i = 1; i < rows.length; i++) {
    if (safeString(rows[i][0]) === safeString(telegramId)) {
      await updateRowCells(SHEET_STATES, i + 1, 'B', [STATES.NONE, '', nowIso()]);
      logEvent('STATE', { telegramId, state: 'none' });
      return;
    }
  }
}

// ===== TELEGRAM =====
const processedUpdates = new Set();
const processedCallbacks = new Set();

const bot = new TelegramBot(BOT_TOKEN, {
  polling: false,
  request: {
    timeout: 30000,
  },
});

const originalBotProcessUpdate = bot.processUpdate.bind(bot);
bot.processUpdate = (update) => {
  if (update?.message) update.message.__doncar_update_id = update.update_id;
  if (update?.callback_query) update.callback_query.__doncar_update_id = update.update_id;
  return originalBotProcessUpdate(update);
};

function clientKeyboard() {
  return {
    keyboard: [
      [{ text: BUTTONS.MY_CARD }, { text: BUTTONS.MY_BALANCE }],
      [{ text: BUTTONS.USE_BONUSES }],
      [{ text: BUTTONS.CONTACT_MANAGER }],
    ],
    resize_keyboard: true,
  };
}

function contactKeyboard() {
  return {
    keyboard: [[{ text: 'Отправить телефон', request_contact: true }]],
    resize_keyboard: true,
    one_time_keyboard: true,
  };
}

function adminKeyboard() {
  return {
    keyboard: [
      [{ text: BUTTONS.FIND_CLIENT }, { text: BUTTONS.CLIENT_BALANCE }],
      [{ text: BUTTONS.CLIENT_HISTORY }, { text: BUTTONS.ACCRUE_BONUSES }],
      [{ text: BUTTONS.MANUAL_ACCRUAL }, { text: BUTTONS.MANUAL_REDEEM }],
      [{ text: BUTTONS.REDEEM_REQUESTS }],
    ],
    resize_keyboard: true,
  };
}

function adminClientActionsKeyboard() {
  return {
    keyboard: [
      [{ text: BUTTONS.CLIENT_HISTORY }, { text: BUTTONS.MANUAL_ACCRUAL }],
      [{ text: BUTTONS.MANUAL_REDEEM }, { text: BUTTONS.ADMIN_MENU }],
    ],
    resize_keyboard: true,
  };
}

async function sendMessage(chatId, text, replyMarkup) {
  const opts = {};
  if (replyMarkup) opts.reply_markup = replyMarkup;
  return await bot.sendMessage(chatId, text, opts);
}

async function notifyAdminsAboutRequest(requestId) {
  const request = await getRequestById(requestId);
  if (!request) return;
  const client = await getClientByTelegramId(request.telegram_id);
  const requestText = `Новая заявка на списание\n${formatRequestCard(request, client)}`;

  for (const adminId of adminIds) {
    await sendMessage(
      adminId,
      requestText,
      {
        inline_keyboard: [[
          { text: 'Подтвердить', callback_data: `${CALLBACKS.APPROVE_REQUEST}${request.request_id}` },
          { text: 'Отклонить', callback_data: `${CALLBACKS.REJECT_REQUEST}${request.request_id}` },
        ]],
      }
    );
  }
}

// ===== FLOW =====
async function sendAdminMenu(chatId, telegramId) {
  await clearState(telegramId);
  return sendMessage(chatId, TEXT.ADMIN_MENU, adminKeyboard());
}

async function findClientOrSendMessage(chatId, identifier) {
  const client = await findClientByIdentifier(identifier);
  if (!client) {
    await sendMessage(chatId, TEXT.CLIENT_NOT_FOUND);
    return null;
  }

  return client;
}

async function handleStart(chatId, telegramId) {
  const client = await getClientByTelegramId(telegramId);

  if (client) {
    await clearState(telegramId);
    return sendMessage(
      chatId,
      `Добро пожаловать в Doncar Club.\nВаш баланс: ${round2(client.bonus_balance)} бонусов`,
      clientKeyboard()
    );
  }

  await setState(telegramId, STATES.WAITING_NAME, {});
  return sendMessage(chatId, 'Добро пожаловать в Doncar Club.\n\nДля регистрации отправьте ваше имя.');
}

async function handleContact(chatId, telegramId, contact) {
  const state = await getState(telegramId);

  if (state !== STATES.WAITING_PHONE_CONTACT) {
    return sendMessage(chatId, TEXT.CONTACT_NOT_REQUESTED);
  }

  if (safeString(contact.user_id) !== telegramId) {
    return sendMessage(chatId, TEXT.SEND_OWN_CONTACT);
  }

  const phone = normalizePhone(contact.phone_number);
  if (!phone) {
    return sendMessage(chatId, TEXT.INVALID_PHONE);
  }

  const temp = await parseStateData(telegramId);
  const name = temp.name || 'Без имени';

  await upsertClient(telegramId, name, phone);
  await clearState(telegramId);

  return sendMessage(chatId, TEXT.REGISTER_COMPLETE, clientKeyboard());
}

async function handleMyCard(chatId, telegramId) {
  const client = await getClientByTelegramId(telegramId);
  if (!client) return sendMessage(chatId, TEXT.REGISTER_FIRST);

  return sendMessage(
    chatId,
    `Карта Doncar Club\n\nИмя: ${client.name}\nТелефон: ${formatPhoneDisplay(client.phone)}\nБаланс: ${round2(client.bonus_balance)} бонусов`,
    clientKeyboard()
  );
}

async function handleMyBalance(chatId, telegramId) {
  const client = await getClientByTelegramId(telegramId);
  if (!client) return sendMessage(chatId, TEXT.REGISTER_FIRST);

  return sendMessage(chatId, `Ваш баланс: ${round2(client.bonus_balance)} бонусов`, clientKeyboard());
}

async function handleUseBonuses(chatId, telegramId) {
  const client = await getClientByTelegramId(telegramId);
  if (!client) return sendMessage(chatId, TEXT.REGISTER_FIRST);

  await setState(telegramId, STATES.CLIENT_WAITING_RENTAL_AMOUNT_FOR_REDEEM, {});
  return sendMessage(
    chatId,
    'Введите сумму новой аренды. Бот рассчитает, сколько бонусов можно использовать.\nЛимит: до 10% от суммы аренды.'
  );
}

async function sendPendingRequests(chatId) {
  const requests = (await getPendingRequests())
    .sort((a, b) => dateTimeToTimestamp(b.created_at) - dateTimeToTimestamp(a.created_at))
    .slice(0, 10);

  if (!requests.length) {
    return sendMessage(chatId, TEXT.NO_PENDING_REQUESTS);
  }

  for (const req of requests) {
    const client = await getClientByTelegramId(req.telegram_id);
    await sendMessage(
      chatId,
      formatRequestCard(req, client),
      {
        inline_keyboard: [[
          { text: 'Подтвердить', callback_data: `${CALLBACKS.APPROVE_REQUEST}${req.request_id}` },
          { text: 'Отклонить', callback_data: `${CALLBACKS.REJECT_REQUEST}${req.request_id}` },
        ]],
      }
    );
  }
}

async function setAdminClientFocus(adminTelegramId, clientTelegramId) {
  await setState(adminTelegramId, STATES.ADMIN_CLIENT_ACTIONS, {
    target_telegram_id: clientTelegramId,
  });
}

async function sendFoundClientCard(chatId, identifier, replyMarkup) {
  const client = await findClientByIdentifier(identifier);
  if (!client) return sendMessage(chatId, TEXT.CLIENT_NOT_FOUND);

  return sendMessage(chatId, formatFoundClientMessage(client), replyMarkup);
}

async function sendClientSummary(chatId, identifier, replyMarkup) {
  const client = await findClientByIdentifier(identifier);
  if (!client) return sendMessage(chatId, TEXT.CLIENT_NOT_FOUND);

  const history = await getClientHistory(client.telegram_id, 10);

  return sendMessage(chatId, formatClientSummaryMessage(client, history), replyMarkup);
}

async function sendClientBalanceInfo(chatId, identifier, replyMarkup) {
  const client = await findClientByIdentifier(identifier);
  if (!client) return sendMessage(chatId, TEXT.CLIENT_NOT_FOUND);

  return sendMessage(chatId, formatBalanceMessage(client), replyMarkup);
}

async function approveRequest(chatId, adminTelegramId, requestId) {
  const request = await getRequestById(requestId);
  if (!request) return sendMessage(chatId, TEXT.REQUEST_NOT_FOUND);
  if (request.status !== REQUEST_STATUS.PENDING) {
    return sendMessage(chatId, formatRequestProcessedMessage(request.status));
  }

  const client = await getClientByTelegramId(request.telegram_id);
  if (!client) return sendMessage(chatId, TEXT.CLIENT_NOT_FOUND);

  const duplicateKey = buildDuplicateKey(
    OPERATION_TYPE.REDEEM,
    request.telegram_id,
    request.rental_amount,
    request.rental_datetime,
    request.rental_id
  );

  if (await historyDuplicateKeyExists(duplicateKey)) {
    await markRequestStatus(requestId, REQUEST_STATUS.DUPLICATE, adminTelegramId);
    return sendMessage(chatId, TEXT.DUPLICATE_REDEEM);
  }

  const allowed = round2(Math.min(
    client.bonus_balance,
    request.max_allowed_bonus,
    request.requested_bonus
  ));

  if (allowed <= 0) {
    await markRequestStatus(requestId, REQUEST_STATUS.REJECTED, adminTelegramId);
    return sendMessage(chatId, TEXT.ZERO_REDEEM);
  }

  const operationId = buildOperationId(OPERATION_TYPE.REDEEM);
  const nextBalance = await changeBonusBalance(request.telegram_id, -allowed);
  await addHistory({
    operation_id: operationId,
    telegram_id: request.telegram_id,
    type: OPERATION_TYPE.REDEEM,
    amount: allowed,
    comment: `Списание на аренду ${request.rental_amount} ₽`,
    admin_id: adminTelegramId,
    rental_id: request.rental_id,
    rental_datetime: request.rental_datetime,
    duplicate_key: duplicateKey,
  });

  await markRequestStatus(requestId, REQUEST_STATUS.APPROVED, adminTelegramId);
  logEvent('REQUEST', { action: 'approved', requestId, adminId: adminTelegramId, telegramId: request.telegram_id });
  logEvent('REDEEM', {
    telegramId: request.telegram_id,
    adminId: adminTelegramId,
    amount: allowed,
    balance: nextBalance,
    requestId,
  });

  await sendMessage(chatId, `Списание подтверждено.\n${formatRedeemMessage(allowed, nextBalance)}`);
  return sendMessage(
    request.telegram_id,
    `Ваш запрос подтверждён.\nСписано: ${allowed} бонусов\nТекущий баланс: ${nextBalance}`,
    clientKeyboard()
  );
}

async function rejectRequest(chatId, adminTelegramId, requestId) {
  const request = await getRequestById(requestId);
  if (!request) return sendMessage(chatId, TEXT.REQUEST_NOT_FOUND);
  if (request.status !== REQUEST_STATUS.PENDING) {
    return sendMessage(chatId, formatRequestProcessedMessage(request.status));
  }

  await markRequestStatus(requestId, REQUEST_STATUS.REJECTED, adminTelegramId);
  logEvent('REQUEST', { action: 'rejected', requestId, adminId: adminTelegramId, telegramId: request.telegram_id });
  await sendMessage(chatId, formatRequestRejectedMessage(requestId));
  return sendMessage(
    request.telegram_id,
    TEXT.REQUEST_REJECTED_CLIENT,
    clientKeyboard()
  );
}

async function buildMessageContext(message) {
  const chatId = safeString(message.chat?.id);
  const telegramId = safeString(message.from?.id);
  const text = safeString(message.text).trim();
  const state = await getState(telegramId);
  const isAdminUser = isAdmin(telegramId);
  const stateData = isAdminUser ? await parseStateData(telegramId) : {};
  const focusedClientId =
    isAdminUser && state === STATES.ADMIN_CLIENT_ACTIONS
      ? safeString(stateData.target_telegram_id)
      : '';

  return {
    message,
    chatId,
    telegramId,
    text,
    state,
    stateData,
    isAdminUser,
    focusedClientId,
  };
}

async function handleCallbackQuery(callbackQuery) {
  const callbackId = callbackQuery.id;
  const chatId = safeString(callbackQuery.message?.chat?.id);
  const adminTelegramId = safeString(callbackQuery.from?.id);
  const data = callbackQuery.data || '';

  try {
    logEvent('UPDATE', { type: 'callback', telegramId: adminTelegramId, data, callbackId });

    if (processedCallbacks.has(callbackId)) {
      logEvent('CALLBACK_DUPLICATE', { callbackId, telegramId: adminTelegramId });
      try {
        await bot.answerCallbackQuery(callbackId);
      } catch (answerError) {
        logError('ERROR', answerError, { scope: 'callback_duplicate_answer', callbackId, telegramId: adminTelegramId });
      }
      return null;
    }

    rememberProcessed(processedCallbacks, callbackId);
    await bot.answerCallbackQuery(callbackId);

    if (!isAdmin(adminTelegramId)) {
      return sendMessage(chatId, TEXT.ACCESS_DENIED_ACTION);
    }

    if (data.startsWith(CALLBACKS.APPROVE_REQUEST)) {
      return approveRequest(chatId, adminTelegramId, data.slice(CALLBACKS.APPROVE_REQUEST.length));
    }

    if (data.startsWith(CALLBACKS.REJECT_REQUEST)) {
      return rejectRequest(chatId, adminTelegramId, data.slice(CALLBACKS.REJECT_REQUEST.length));
    }

    return null;
  } catch (error) {
    logError('ERROR', error, {
      scope: 'callback_query',
      callbackId,
      telegramId: adminTelegramId,
      data,
    });
    return sendMessage(chatId, TEXT.CALLBACK_ACTION_ERROR);
  }
}

async function handleCommand(context) {
  const { chatId, telegramId, text } = context;

  if (text === '/start') {
    return handleStart(chatId, telegramId);
  }

  if (text === '/admin') {
    if (!context.isAdminUser) return sendMessage(chatId, TEXT.ACCESS_DENIED_ADMIN);
    return sendAdminMenu(chatId, telegramId);
  }

  const clientCommandMatch = text.match(/^\/client(?:@\w+)?(?:\s+(.+))?$/);
  if (!clientCommandMatch) return null;
  if (!context.isAdminUser) return sendMessage(chatId, TEXT.ACCESS_DENIED_ADMIN);

  const identifier = safeString(clientCommandMatch[1]).trim();
  if (!identifier) {
    return sendMessage(chatId, 'Использование:\n/client <telegram_id_или_телефон>');
  }

  const client = await findClientOrSendMessage(chatId, identifier);
  if (!client) return true;

  await setAdminClientFocus(telegramId, client.telegram_id);
  return sendClientSummary(chatId, client.telegram_id, adminClientActionsKeyboard());
}

async function handleClientMenuAction(context) {
  const { chatId, telegramId, text } = context;

  if (text === BUTTONS.MY_CARD) return handleMyCard(chatId, telegramId);
  if (text === BUTTONS.MY_BALANCE) return handleMyBalance(chatId, telegramId);
  if (text === BUTTONS.USE_BONUSES) return handleUseBonuses(chatId, telegramId);
  if (text === BUTTONS.CONTACT_MANAGER) {
    return sendMessage(chatId, `Связаться с менеджером: @${MANAGER_USERNAME}`, clientKeyboard());
  }

  return null;
}

async function handleAdminAction(context) {
  const { chatId, telegramId, text, isAdminUser, focusedClientId } = context;
  if (!isAdminUser) return null;

  if (text === BUTTONS.ADMIN_MENU) {
    return sendAdminMenu(chatId, telegramId);
  }

  if (focusedClientId && text === BUTTONS.CLIENT_HISTORY) {
    await setAdminClientFocus(telegramId, focusedClientId);
    return sendClientSummary(chatId, focusedClientId, adminClientActionsKeyboard());
  }

  if (focusedClientId && text === BUTTONS.MANUAL_ACCRUAL) {
    await setState(telegramId, STATES.ADMIN_WAITING_MANUAL_ACCRUAL_AMOUNT, {
      target_telegram_id: focusedClientId,
    });
    return sendMessage(chatId, TEXT.WAIT_MANUAL_ACCRUAL_AMOUNT);
  }

  if (focusedClientId && text === BUTTONS.MANUAL_REDEEM) {
    await setState(telegramId, STATES.ADMIN_WAITING_MANUAL_REDEEM_AMOUNT, {
      target_telegram_id: focusedClientId,
    });
    return sendMessage(chatId, TEXT.WAIT_MANUAL_REDEEM_AMOUNT);
  }

  if (text === BUTTONS.FIND_CLIENT) {
    await setState(telegramId, STATES.ADMIN_WAITING_FIND_IDENTIFIER, {});
    return sendMessage(chatId, TEXT.WAIT_IDENTIFIER);
  }

  if (text === BUTTONS.CLIENT_BALANCE) {
    await setState(telegramId, STATES.ADMIN_WAITING_BALANCE_IDENTIFIER, {});
    return sendMessage(chatId, TEXT.WAIT_IDENTIFIER);
  }

  if (text === BUTTONS.CLIENT_HISTORY) {
    await setState(telegramId, STATES.ADMIN_WAITING_HISTORY_IDENTIFIER, {});
    return sendMessage(chatId, TEXT.WAIT_IDENTIFIER);
  }

  if (text === BUTTONS.ACCRUE_BONUSES) {
    await setState(telegramId, STATES.ADMIN_WAITING_CLIENT_IDENTIFIER, {});
    return sendMessage(chatId, TEXT.WAIT_IDENTIFIER);
  }

  if (text === BUTTONS.MANUAL_ACCRUAL) {
    await setState(telegramId, STATES.ADMIN_WAITING_MANUAL_ACCRUAL_IDENTIFIER, {});
    return sendMessage(chatId, TEXT.WAIT_IDENTIFIER);
  }

  if (text === BUTTONS.MANUAL_REDEEM) {
    await setState(telegramId, STATES.ADMIN_WAITING_MANUAL_REDEEM_IDENTIFIER, {});
    return sendMessage(chatId, TEXT.WAIT_IDENTIFIER);
  }

  if (text === BUTTONS.REDEEM_REQUESTS) {
    return sendPendingRequests(chatId);
  }

  return null;
}

async function handleWaitingNameState(context) {
  const { chatId, telegramId, text } = context;

  if (!text || text.length < 2) {
    return sendMessage(chatId, TEXT.INVALID_NAME);
  }

  await setState(telegramId, STATES.WAITING_PHONE_CONTACT, { name: text });
  return sendMessage(chatId, TEXT.SEND_CONTACT_PROMPT, contactKeyboard());
}

async function handleClientWaitingRentalAmountState(context) {
  const { chatId, telegramId, text } = context;
  const amount = parseMoney(text);
  if (!amount || amount <= 0) return sendMessage(chatId, TEXT.INVALID_RENTAL_AMOUNT);

  const client = await getClientByTelegramId(telegramId);
  if (!client) {
    await clearState(telegramId);
    return sendMessage(chatId, TEXT.REGISTER_FIRST);
  }

  const maxAllowed = round2(amount * 0.10);
  const requested = round2(Math.min(client.bonus_balance, maxAllowed));

  if (requested <= 0) {
    await clearState(telegramId);
    return sendMessage(chatId, formatNoRedeemAvailableMessage(client), clientKeyboard());
  }

  await setState(telegramId, STATES.CLIENT_WAITING_REDEEM_DATETIME, {
    rental_amount: amount,
    max_allowed_bonus: maxAllowed,
    requested_bonus: requested,
  });

  return sendMessage(chatId, formatAdminRentalCalculation(amount, client, maxAllowed, requested));
}

async function handleClientWaitingRedeemDatetimeState(context) {
  const { chatId, telegramId, text } = context;
  const rentalDateTime = normalizeDateTime(text);
  if (!rentalDateTime) return sendMessage(chatId, TEXT.DATE_PROMPT);

  const temp = await parseStateData(telegramId);
  temp.rental_datetime = rentalDateTime;
  await setState(telegramId, STATES.CLIENT_WAITING_REDEEM_RENTAL_ID, temp);

  return sendMessage(chatId, TEXT.REQUEST_RENTAL_ID_PROMPT);
}

async function handleClientWaitingRedeemRentalIdState(context) {
  const { chatId, telegramId, text } = context;
  const temp = await parseStateData(telegramId);
  const client = await getClientByTelegramId(telegramId);

  if (!client) {
    await clearState(telegramId);
    return sendMessage(chatId, TEXT.CLIENT_NOT_FOUND_REGISTER);
  }

  const rentalId = text === '-' ? '' : text;

  if (await requestLooksDuplicate(telegramId, temp.rental_amount, temp.rental_datetime, rentalId)) {
    await clearState(telegramId);
    return sendMessage(chatId, TEXT.SAME_REQUEST_EXISTS, clientKeyboard());
  }

  const requestId = buildRequestId(telegramId);

  await addRequest({
    request_id: requestId,
    telegram_id: telegramId,
    phone: client.phone,
    rental_amount: temp.rental_amount,
    max_allowed_bonus: temp.max_allowed_bonus,
    requested_bonus: temp.requested_bonus,
    status: REQUEST_STATUS.PENDING,
    admin_id: '',
    rental_id: rentalId,
    rental_datetime: temp.rental_datetime,
  });

  logEvent('REQUEST', {
    action: 'created',
    telegramId,
    requestId,
    requestedBonus: temp.requested_bonus,
    rentalAmount: temp.rental_amount,
  });

  await clearState(telegramId);
  await sendMessage(chatId, formatRequestCreatedMessage(requestId, temp.requested_bonus), clientKeyboard());
  return notifyAdminsAboutRequest(requestId);
}

async function handleAdminWaitingFindIdentifierState(context) {
  const client = await findClientOrSendMessage(context.chatId, context.text);
  if (!client) return true;

  await setAdminClientFocus(context.telegramId, client.telegram_id);
  return sendFoundClientCard(context.chatId, client.telegram_id, adminClientActionsKeyboard());
}

async function handleAdminWaitingHistoryIdentifierState(context) {
  const client = await findClientOrSendMessage(context.chatId, context.text);
  if (!client) return true;

  await setAdminClientFocus(context.telegramId, client.telegram_id);
  return sendClientSummary(context.chatId, client.telegram_id, adminClientActionsKeyboard());
}

async function handleAdminWaitingBalanceIdentifierState(context) {
  const client = await findClientOrSendMessage(context.chatId, context.text);
  if (!client) return true;

  await setAdminClientFocus(context.telegramId, client.telegram_id);
  return sendClientBalanceInfo(context.chatId, client.telegram_id, adminClientActionsKeyboard());
}

async function handleAdminWaitingClientIdentifierState(context) {
  const client = await findClientOrSendMessage(context.chatId, context.text);
  if (!client) return true;

  await setState(context.telegramId, STATES.ADMIN_WAITING_ACCRUAL_AMOUNT, {
    target_telegram_id: client.telegram_id,
  });

  return sendMessage(context.chatId, formatClientPromptMessage(client, TEXT.WAIT_RENTAL_AMOUNT));
}

async function handleAdminWaitingManualAccrualIdentifierState(context) {
  const client = await findClientOrSendMessage(context.chatId, context.text);
  if (!client) return true;

  await setState(context.telegramId, STATES.ADMIN_WAITING_MANUAL_ACCRUAL_AMOUNT, {
    target_telegram_id: client.telegram_id,
  });

  return sendMessage(context.chatId, formatClientPromptMessage(client, TEXT.WAIT_MANUAL_ACCRUAL_AMOUNT));
}

async function handleAdminWaitingManualAccrualAmountState(context) {
  const amount = parseMoney(context.text);
  if (!amount || amount <= 0) return sendMessage(context.chatId, TEXT.INVALID_AMOUNT);

  const temp = await parseStateData(context.telegramId);
  temp.bonus_amount = amount;
  await setState(context.telegramId, STATES.ADMIN_WAITING_MANUAL_ACCRUAL_COMMENT, temp);

  return sendMessage(context.chatId, TEXT.WAIT_COMMENT);
}

async function handleAdminWaitingManualAccrualCommentState(context) {
  const temp = await parseStateData(context.telegramId);
  const targetTelegramId = safeString(temp.target_telegram_id);
  const client = await getClientByTelegramId(targetTelegramId);

  if (!client) {
    await setState(context.telegramId, STATES.ADMIN_WAITING_MANUAL_ACCRUAL_IDENTIFIER, {});
    return sendMessage(context.chatId, TEXT.CLIENT_NOT_FOUND);
  }

  const comment = context.text.trim();
  if (!comment) return sendMessage(context.chatId, TEXT.COMMENT_REQUIRED);

  const duplicateKey = buildManualAccrualDuplicateKey(targetTelegramId, temp.bonus_amount, comment);
  if (await historyDuplicateKeyExists(duplicateKey)) {
    return sendMessage(context.chatId, TEXT.DUPLICATE_MANUAL_OPERATION);
  }

  const operationId = buildOperationId(OPERATION_TYPE.MANUAL_ACCRUAL);
  const rentalDateTime = currentDisplayDateTime();
  const nextBalance = await changeBonusBalance(targetTelegramId, temp.bonus_amount);

  await addHistory({
    operation_id: operationId,
    telegram_id: targetTelegramId,
    type: OPERATION_TYPE.MANUAL_ACCRUAL,
    amount: temp.bonus_amount,
    comment,
    admin_id: context.telegramId,
    rental_datetime: rentalDateTime,
    duplicate_key: duplicateKey,
  });

  logEvent('MANUAL_ACCRUAL', {
    telegramId: targetTelegramId,
    adminId: context.telegramId,
    amount: temp.bonus_amount,
    balance: nextBalance,
  });

  await setAdminClientFocus(context.telegramId, targetTelegramId);
  await sendMessage(context.chatId, formatAccrualMessage(temp.bonus_amount, nextBalance));
  await sendFoundClientCard(context.chatId, targetTelegramId, adminClientActionsKeyboard());

  return sendMessage(
    targetTelegramId,
    `Вам начислено ${temp.bonus_amount} бонусов.\nТекущий баланс: ${nextBalance} бонусов.`,
    clientKeyboard()
  );
}

async function handleAdminWaitingManualRedeemIdentifierState(context) {
  const client = await findClientOrSendMessage(context.chatId, context.text);
  if (!client) return true;

  await setState(context.telegramId, STATES.ADMIN_WAITING_MANUAL_REDEEM_AMOUNT, {
    target_telegram_id: client.telegram_id,
  });

  return sendMessage(context.chatId, formatClientPromptMessage(client, TEXT.WAIT_MANUAL_REDEEM_AMOUNT));
}

async function handleAdminWaitingManualRedeemAmountState(context) {
  const amount = parseMoney(context.text);
  if (!amount || amount <= 0) return sendMessage(context.chatId, TEXT.INVALID_AMOUNT);

  const temp = await parseStateData(context.telegramId);
  temp.bonus_amount = amount;
  await setState(context.telegramId, STATES.ADMIN_WAITING_MANUAL_REDEEM_COMMENT, temp);

  return sendMessage(context.chatId, TEXT.WAIT_COMMENT);
}

async function handleAdminWaitingManualRedeemCommentState(context) {
  const temp = await parseStateData(context.telegramId);
  const targetTelegramId = safeString(temp.target_telegram_id);
  const client = await getClientByTelegramId(targetTelegramId);

  if (!client) {
    await setState(context.telegramId, STATES.ADMIN_WAITING_MANUAL_REDEEM_IDENTIFIER, {});
    return sendMessage(context.chatId, TEXT.CLIENT_NOT_FOUND);
  }

  const comment = context.text.trim();
  if (!comment) return sendMessage(context.chatId, TEXT.COMMENT_REQUIRED);

  if (client.bonus_balance < safeNumber(temp.bonus_amount)) {
    await setState(context.telegramId, STATES.ADMIN_WAITING_MANUAL_REDEEM_AMOUNT, {
      target_telegram_id: targetTelegramId,
    });
    return sendMessage(context.chatId, TEXT.NOT_ENOUGH_BONUSES);
  }

  const duplicateKey = buildManualRedeemDuplicateKey(targetTelegramId, temp.bonus_amount, comment);
  if (await historyDuplicateKeyExists(duplicateKey)) {
    return sendMessage(context.chatId, TEXT.DUPLICATE_MANUAL_OPERATION);
  }

  const operationId = buildOperationId(OPERATION_TYPE.MANUAL_REDEEM);
  const rentalDateTime = currentDisplayDateTime();
  const nextBalance = await changeBonusBalance(targetTelegramId, -temp.bonus_amount);

  await addHistory({
    operation_id: operationId,
    telegram_id: targetTelegramId,
    type: OPERATION_TYPE.MANUAL_REDEEM,
    amount: temp.bonus_amount,
    comment,
    admin_id: context.telegramId,
    rental_datetime: rentalDateTime,
    duplicate_key: duplicateKey,
  });

  logEvent('MANUAL_REDEEM', {
    telegramId: targetTelegramId,
    adminId: context.telegramId,
    amount: temp.bonus_amount,
    balance: nextBalance,
  });

  await setAdminClientFocus(context.telegramId, targetTelegramId);
  await sendMessage(context.chatId, formatRedeemMessage(temp.bonus_amount, nextBalance));
  await sendFoundClientCard(context.chatId, targetTelegramId, adminClientActionsKeyboard());

  return sendMessage(
    targetTelegramId,
    `С вашего бонусного баланса списано ${temp.bonus_amount} бонусов.\nТекущий баланс: ${nextBalance} бонусов.`,
    clientKeyboard()
  );
}

async function handleAdminWaitingAccrualAmountState(context) {
  const amount = parseMoney(context.text);
  if (!amount || amount <= 0) return sendMessage(context.chatId, TEXT.INVALID_RENTAL_AMOUNT);

  const temp = await parseStateData(context.telegramId);
  temp.rental_amount = amount;
  await setState(context.telegramId, STATES.ADMIN_WAITING_ACCRUAL_DATETIME, temp);

  return sendMessage(context.chatId, TEXT.DATE_PROMPT);
}

async function handleAdminWaitingAccrualDatetimeState(context) {
  const rentalDateTime = normalizeDateTime(context.text);
  if (!rentalDateTime) return sendMessage(context.chatId, TEXT.DATE_PROMPT);

  const temp = await parseStateData(context.telegramId);
  temp.rental_datetime = rentalDateTime;
  await setState(context.telegramId, STATES.ADMIN_WAITING_ACCRUAL_RENTAL_ID, temp);

  return sendMessage(context.chatId, TEXT.REQUEST_RENTAL_ID_PROMPT_ADMIN);
}

async function handleAdminWaitingAccrualRentalIdState(context) {
  const temp = await parseStateData(context.telegramId);
  const rentalId = context.text === '-' ? '' : context.text;
  const targetTelegramId = safeString(temp.target_telegram_id);
  const client = await getClientByTelegramId(targetTelegramId);

  if (!client) {
    await clearState(context.telegramId);
    return sendMessage(context.chatId, TEXT.CLIENT_NOT_FOUND_RESTART);
  }

  const duplicateKey = buildDuplicateKey(
    OPERATION_TYPE.ACCRUAL,
    targetTelegramId,
    temp.rental_amount,
    temp.rental_datetime,
    rentalId
  );

  if (await historyDuplicateKeyExists(duplicateKey)) {
    await clearState(context.telegramId);
    return sendMessage(context.chatId, TEXT.DUPLICATE_ACCRUAL);
  }

  const bonus = round2(temp.rental_amount * 0.05);
  const operationId = buildOperationId(OPERATION_TYPE.ACCRUAL);
  const nextBalance = await changeBonusBalance(targetTelegramId, bonus);
  await addHistory({
    operation_id: operationId,
    telegram_id: targetTelegramId,
    type: OPERATION_TYPE.ACCRUAL,
    amount: bonus,
    comment: `5% от аренды ${temp.rental_amount} ₽`,
    admin_id: context.telegramId,
    rental_id: rentalId,
    rental_datetime: temp.rental_datetime,
    duplicate_key: duplicateKey,
  });

  logEvent('ACCRUAL', {
    telegramId: targetTelegramId,
    adminId: context.telegramId,
    amount: bonus,
    balance: nextBalance,
  });

  await setAdminClientFocus(context.telegramId, targetTelegramId);
  await sendMessage(context.chatId, formatAccrualMessage(bonus, nextBalance));
  await sendClientSummary(context.chatId, targetTelegramId, adminClientActionsKeyboard());

  return sendMessage(
    targetTelegramId,
    `Вам начислено ${bonus} бонусов.\nТекущий баланс: ${nextBalance} бонусов.`,
    clientKeyboard()
  );
}

const STATE_HANDLERS = {
  [STATES.WAITING_NAME]: handleWaitingNameState,
  [STATES.CLIENT_WAITING_RENTAL_AMOUNT_FOR_REDEEM]: handleClientWaitingRentalAmountState,
  [STATES.CLIENT_WAITING_REDEEM_DATETIME]: handleClientWaitingRedeemDatetimeState,
  [STATES.CLIENT_WAITING_REDEEM_RENTAL_ID]: handleClientWaitingRedeemRentalIdState,
  [STATES.ADMIN_WAITING_FIND_IDENTIFIER]: handleAdminWaitingFindIdentifierState,
  [STATES.ADMIN_WAITING_HISTORY_IDENTIFIER]: handleAdminWaitingHistoryIdentifierState,
  [STATES.ADMIN_WAITING_BALANCE_IDENTIFIER]: handleAdminWaitingBalanceIdentifierState,
  [STATES.ADMIN_WAITING_CLIENT_IDENTIFIER]: handleAdminWaitingClientIdentifierState,
  [STATES.ADMIN_WAITING_ACCRUAL_AMOUNT]: handleAdminWaitingAccrualAmountState,
  [STATES.ADMIN_WAITING_ACCRUAL_DATETIME]: handleAdminWaitingAccrualDatetimeState,
  [STATES.ADMIN_WAITING_ACCRUAL_RENTAL_ID]: handleAdminWaitingAccrualRentalIdState,
  [STATES.ADMIN_WAITING_MANUAL_ACCRUAL_IDENTIFIER]: handleAdminWaitingManualAccrualIdentifierState,
  [STATES.ADMIN_WAITING_MANUAL_ACCRUAL_AMOUNT]: handleAdminWaitingManualAccrualAmountState,
  [STATES.ADMIN_WAITING_MANUAL_ACCRUAL_COMMENT]: handleAdminWaitingManualAccrualCommentState,
  [STATES.ADMIN_WAITING_MANUAL_REDEEM_IDENTIFIER]: handleAdminWaitingManualRedeemIdentifierState,
  [STATES.ADMIN_WAITING_MANUAL_REDEEM_AMOUNT]: handleAdminWaitingManualRedeemAmountState,
  [STATES.ADMIN_WAITING_MANUAL_REDEEM_COMMENT]: handleAdminWaitingManualRedeemCommentState,
  [STATES.ADMIN_WAITING_MANUAL_CLIENT_IDENTIFIER]: handleAdminWaitingManualAccrualIdentifierState,
  [STATES.ADMIN_WAITING_MANUAL_AMOUNT]: handleAdminWaitingManualAccrualAmountState,
  [STATES.ADMIN_WAITING_MANUAL_COMMENT]: handleAdminWaitingManualAccrualCommentState,
};

async function handleStateFlow(context) {
  const handler = STATE_HANDLERS[context.state];
  if (!handler) return null;
  if (safeString(context.state).startsWith('admin_') && !context.isAdminUser) return null;
  return handler(context);
}

async function handleMessageUpdate(message) {
  const context = await buildMessageContext(message);

  logEvent('UPDATE', {
    type: 'message',
    telegramId: context.telegramId,
    text: context.text || '[non-text]',
    state: context.state || STATES.NONE,
  });

  if (message.contact) return handleContact(context.chatId, context.telegramId, message.contact);

  const handlers = [
    handleCommand,
    handleClientMenuAction,
    handleAdminAction,
    handleStateFlow,
  ];

  for (const handler of handlers) {
    const result = await handler(context);
    if (result) return result;
  }

  if (context.isAdminUser) {
    return sendMessage(
      context.chatId,
      TEXT.ACTION_PROMPT,
      context.focusedClientId ? adminClientActionsKeyboard() : adminKeyboard()
    );
  }

  return sendMessage(context.chatId, TEXT.UNKNOWN_COMMAND, clientKeyboard());
}

async function processUpdate(update) {
  const updateId = safeString(update?.update_id);
  const chatId = safeString(update?.callback_query?.message?.chat?.id || update?.message?.chat?.id);
  const telegramId = safeString(update?.callback_query?.from?.id || update?.message?.from?.id);

  try {
    if (updateId && processedUpdates.has(updateId)) {
      logEvent('UPDATE_DUPLICATE', { updateId, telegramId });
      return null;
    }

    rememberProcessed(processedUpdates, updateId);

    if (update.callback_query) return handleCallbackQuery(update.callback_query);
    if (update.message) return handleMessageUpdate(update.message);
    return null;
  } catch (error) {
    logError('ERROR', error, { scope: 'processUpdate', chatId, telegramId });

    if (chatId) {
      try {
        await sendMessage(chatId, TEXT.SERVICE_UNAVAILABLE);
      } catch (sendError) {
        logError('ERROR', sendError, { scope: 'serviceUnavailableResponse', chatId, telegramId });
      }
    }

    return null;
  }
}

function bindPollingHandlers() {
  bot.removeAllListeners('message');
  bot.removeAllListeners('callback_query');

  bot.on('message', async (message) => {
    try {
      await processUpdate({
        update_id: message.__doncar_update_id,
        message,
      });
    } catch (error) {
      logError('ERROR', error, { scope: 'polling_message' });
    }
  });

  bot.on('callback_query', async (callback_query) => {
    try {
      await processUpdate({
        update_id: callback_query.__doncar_update_id,
        callback_query,
      });
    } catch (error) {
      logError('ERROR', error, { scope: 'polling_callback' });
    }
  });
}

// ===== SERVER =====
const app = express();
app.use(express.json({ limit: '1mb' }));

app.get('/', (_req, res) => {
  res.status(200).json({
    ok: true,
    app: 'doncar-bot',
    version: APP_VERSION,
  });
});

app.post('/telegram/webhook', async (req, res) => {
  try {
    await processUpdate(req.body);
    res.status(200).send('ok');
  } catch (error) {
    logError('ERROR', error, { scope: 'webhook' });
    res.status(200).send('ok');
  }
});

async function start() {
  console.log('[BOT] Starting Doncar Bot');
  console.log('[BOT] Mode:', USE_WEBHOOK === 'true' ? 'WEBHOOK' : 'POLLING');

  await initSheets();
  bindPollingHandlers();

  if (USE_WEBHOOK === 'true') {
    // Webhook режим используется для production.
    if (!BASE_URL) throw new Error('BASE_URL is required when USE_WEBHOOK=true');
    const webhookUrl = `${BASE_URL.replace(/\/$/, '')}/telegram/webhook`;
    await bot.setWebHook(webhookUrl, { drop_pending_updates: true });
    console.log('[BOT] Webhook set:', webhookUrl);
    logEvent('STARTUP', { mode: 'webhook', webhookUrl });
  } else {
    // Polling режим используется для разработки.
    await bot.deleteWebHook({ drop_pending_updates: true });
    await bot.startPolling({
      interval: 300,
      autoStart: true,
      params: {
        timeout: 30,
      },
    });
    console.log('[BOT] Polling started');
    logEvent('STARTUP', { mode: 'polling' });
  }

  app.listen(safeNumber(PORT), () => {
    logEvent('STARTUP', { port: PORT, version: APP_VERSION });
  });
}

start().catch(err => {
  logError('ERROR', err, { scope: 'startup' });
  process.exit(1);
});
