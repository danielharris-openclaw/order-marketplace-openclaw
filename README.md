# Биржа заказов

Статический сайт с хранением заявок в Google Sheets.

## Google Sheets

Таблица: https://docs.google.com/spreadsheets/d/1kTh8RZaAN_jshluKPJoBMhTYt4ERCPRltnhFc7-zUIQ/edit

Листы:

- `Заказы`: `id`, `createdAt`, `title`, `summary`, `deadline`, `price`, `phone`, `status`
- `Отклики`: `id`, `orderId`, `createdAt`, `name`, `phone`

Статусы заявок:

- `active` - актуальна
- `in_work` - в работе, включается автоматически после первого отклика
- `closed` - закрыта заказчиком после подтверждения телефона

## Уведомления в MAX

В Google Apps Script откройте Project Settings -> Script properties и добавьте:

- `MAX_BOT_TOKEN` - токен бота MAX
- `MAX_CHAT_ID` - ID группы MAX

Тестовая группа MAX:

- `MAX_CHAT_ID=-75195528071310`

Токен не хранится в публичном репозитории. Скрипт читает его только из Script properties.

## Подключение API

1. Открыть https://script.google.com/
2. Создать новый проект.
3. Вставить код из `google-apps-script.js`.
4. Нажать Deploy -> New deployment -> Web app.
5. Execute as: Me.
6. Who has access: Anyone.
7. Скопировать Web app URL.
8. В `index.html` вставить URL в переменную `apiUrl`.
