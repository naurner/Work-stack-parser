# Telegram Channel Parser

Простой парсер для скачивания сообщений из публичных Telegram каналов в Excel. Работает без API ключей через веб-версию Telegram.

## Установка

```bash
pip install beautifulsoup4 aiohttp openpyxl lxml
```

## Использование

Откройте `Parser message.py` и измените настройки в начале файла:

```python
CHANNEL_USERNAME = '@durov'
EXCEL_FILE = 'telegram_messages.xlsx'
PARSE_PERIOD_DAYS = 365
FILTER_KEYWORDS = ''
```

Запустите:

```bash
python "Parser message.py"
```

## Настройки

- `CHANNEL_USERNAME` - имя канала (с @)
- `EXCEL_FILE` - имя выходного файла
- `PARSE_PERIOD_DAYS` - за сколько дней парсить (365 = год)
- `FILTER_KEYWORDS` - ключевые слова для фильтрации (через запятую)

## Примеры

Все сообщения:
```python
CHANNEL_USERNAME = '@durov'
FILTER_KEYWORDS = ''
```

Только с определенными словами:
```python
CHANNEL_USERNAME = '@tech_news'
FILTER_KEYWORDS = 'python,django'
```

## Ограничения

- Работает только с публичными каналами
- Telegram показывает через веб примерно 200-500 последних сообщений
- Нет мониторинга новых сообщений в реальном времени

## Результат

Создается Excel файл с колонками:
- ID сообщения
- Дата
- Время
- Автор
- Текст
- Просмотры
- Ссылка

