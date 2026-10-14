# Telegram Channel Parser

Exports the history of a public Telegram channel to Excel. Works through the public web preview (`t.me/s/<channel>`), so it needs no bot token, no Telegram account and no API keys.

I wrote it to collect job posts from channels like `@findwork` and filter them by keywords.

## How it works

1. Loads `https://t.me/s/<channel>` with `aiohttp` and parses the posts with BeautifulSoup.
2. Pages back through history with `?before=<oldest message id>` until it reaches the start of the period or the message limit.
3. Keeps only posts that contain at least one of the keywords (if keywords are set).
4. Appends rows to an Excel file and skips message IDs that are already there, so re-running the script only adds new posts.

Columns: `ID`, `Дата`, `Время`, `Автор`, `Текст сообщения`, `Просмотры`, `Ссылка`. View counters like `1.2K` are converted to numbers.

## Run

```bash
git clone https://github.com/naurner/Work-stack-parser
cd Work-stack-parser
pip install -r requirements.txt
python channel_parser.py
```

Settings are at the top of `channel_parser.py`:

```python
CHANNEL_USERNAME = '@findwork'          # public channel
EXCEL_FILE = 'telegram_messages.xlsx'
PARSE_PERIOD_DAYS = 365                 # how far back to go
MAX_MESSAGES = 100000
FILTER_KEYWORDS = 'python, django'      # comma-separated, empty = everything
```

## Tests

```bash
pip install pytest
pytest -v
```

## Stack

Python 3.10+ · asyncio · aiohttp · BeautifulSoup4 + lxml · openpyxl · pytest
