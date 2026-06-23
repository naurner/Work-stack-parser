# Telegram Channel Parser

Async scraper that exports messages from public Telegram channels to Excel — without a bot token or Telegram API registration.

## Features

- **No bot token required** — works through the web version of Telegram
- **Async scraping** via `aiohttp` for high performance
- **Flexible filtering** — by keyword, date range, message type
- **Excel export** with `openpyxl` — clean, ready-to-use spreadsheets
- **Configurable** — set channel, period, and filters in a single config block

## Stack

`Python` `aiohttp` `BeautifulSoup4` `openpyxl` `asyncio`

## Quick Start

```bash
git clone https://github.com/naurner/Work-stack-parser
cd Work-stack-parser
pip install -r requirements.txt
```

Edit the config at the top of `Parser message.py`:

```python
CHANNEL_USERNAME = "@channel_name"  # target channel
EXCEL_FILE       = "output.xlsx"
PARSE_PERIOD_DAYS = 30              # how many days back
FILTER_KEYWORDS  = ""               # leave empty to get all messages
```

Then run:

```bash
python "Parser message.py"
```

## Output

An `.xlsx` file with columns: `date`, `time`, `text`, `views`, `link`.

