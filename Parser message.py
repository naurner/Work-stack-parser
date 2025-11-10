import asyncio
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
import aiohttp
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment
import os

CHANNEL_USERNAME = '@findwork'
EXCEL_FILE = 'telegram_messages.xlsx'
PARSE_PERIOD_DAYS = 365
MAX_MESSAGES = 100000
FILTER_KEYWORDS = ''


class TelegramChannelParser:
    def __init__(self, channel_username, excel_file, filter_keywords=''):
        self.channel_username = channel_username.replace('@', '')
        self.excel_file = excel_file
        self.base_url = f"https://t.me/s/{self.channel_username}"
        self.workbook = None
        self.worksheet = None
        self.session = None
        if filter_keywords:
            self.keywords = [kw.strip().lower() for kw in filter_keywords.split(',') if kw.strip()]
        else:
            self.keywords = []

    def matches_filter(self, text):
        if not self.keywords:
            return True

        text_lower = text.lower()
        for keyword in self.keywords:
            if keyword in text_lower:
                return True
        return False

    def init_excel(self):
        if os.path.exists(self.excel_file):
            print(f"Загрузка существующего файла {self.excel_file}")
            self.workbook = load_workbook(self.excel_file)
            self.worksheet = self.workbook.active
        else:
            print(f"Создание нового файла {self.excel_file}")
            self.workbook = Workbook()
            self.worksheet = self.workbook.active
            self.worksheet.title = "Messages"

            headers = ['ID', 'Дата', 'Время', 'Автор', 'Текст сообщения', 'Просмотры', 'Ссылка']
            self.worksheet.append(headers)

            for cell in self.worksheet[1]:
                cell.font = Font(bold=True, size=12)
                cell.alignment = Alignment(horizontal='center', vertical='center')

            self.worksheet.column_dimensions['A'].width = 10
            self.worksheet.column_dimensions['B'].width = 12
            self.worksheet.column_dimensions['C'].width = 10
            self.worksheet.column_dimensions['D'].width = 20
            self.worksheet.column_dimensions['E'].width = 60
            self.worksheet.column_dimensions['F'].width = 12
            self.worksheet.column_dimensions['G'].width = 40

            self.save_excel()

    def save_excel(self):
        self.workbook.save(self.excel_file)
        print(f"Файл {self.excel_file} сохранен")

    def message_exists(self, message_id):
        for row in self.worksheet.iter_rows(min_row=2, max_col=1, values_only=True):
            if row[0] == message_id:
                return True
        return False

    def parse_message_date(self, date_element):
        try:
            if date_element and 'datetime' in date_element.attrs:
                date_str = date_element['datetime']
                dt = datetime.fromisoformat(date_str.replace('Z', '+00:00'))
                return dt.replace(tzinfo=None)
        except:
            pass
        return None

    def parse_views(self, views_element):
        try:
            if views_element:
                views_text = views_element.get_text(strip=True)
                views_text = views_text.replace('K', '000').replace('M', '000000').replace(',', '')
                return int(''.join(filter(str.isdigit, views_text)))
        except:
            pass
        return 0

    async def fetch_page(self, url):
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        async with self.session.get(url, headers=headers) as response:
            if response.status == 200:
                return await response.text()
            else:
                print(f"Ошибка при получении страницы: {response.status}")
                return None

    async def parse_messages_from_page(self, html, start_date):
        soup = BeautifulSoup(html, 'html.parser')
        messages = soup.find_all('div', class_='tgme_widget_message')
        parsed_messages = []


        for msg in messages:
            try:
                message_link = msg.get('data-post', '')
                if not message_link:
                    link_elem = msg.find('a', class_='tgme_widget_message_date')
                    if link_elem and 'href' in link_elem.attrs:
                        message_link = link_elem['href']

                if '/' in message_link:
                    message_id = int(message_link.split('/')[-1])
                else:
                    continue

                if self.message_exists(message_id):
                    continue

                date_element = msg.find('time', class_='datetime')
                if not date_element:
                    date_element = msg.find('time')

                msg_date = self.parse_message_date(date_element)

                if not msg_date:
                    continue

                date_str = msg_date.strftime('%d.%m.%Y')
                time_str = msg_date.strftime('%H:%M:%S')

                author_element = msg.find('div', class_='tgme_widget_message_author')
                if not author_element:
                    author_element = msg.find('a', class_='tgme_widget_message_owner_name')
                author = author_element.get_text(strip=True) if author_element else self.channel_username

                text_element = msg.find('div', class_='tgme_widget_message_text')
                text = ''
                if text_element:
                    text = text_element.get_text('\n', strip=True)

                if not text:
                    photo = msg.find('a', class_='tgme_widget_message_photo_wrap')
                    video = msg.find('video', class_='tgme_widget_message_video')
                    doc = msg.find('div', class_='tgme_widget_message_document')
                    if photo:
                        text = '[Фото]'
                    elif video:
                        text = '[Видео]'
                    elif doc:
                        text = '[Документ]'
                    else:
                        text = '[Медиа]'

                if not self.matches_filter(text):
                    continue


                views_element = msg.find('span', class_='tgme_widget_message_views')
                views = self.parse_views(views_element)

                link = f"https://t.me/{self.channel_username}/{message_id}"

                parsed_messages.append({
                    'id': message_id,
                    'date': date_str,
                    'time': time_str,
                    'author': author,
                    'text': text[:500] if len(text) > 500 else text,
                    'views': views,
                    'link': link,
                    'datetime': msg_date
                })

            except Exception as e:
                print(f"Ошибка парсинга сообщения: {e}")
                import traceback
                traceback.print_exc()
                continue

        return parsed_messages

    async def parse_historical_messages(self, days=365):
        print(f"\n{'='*60}")
        print(f"Начало парсинга сообщений за последние {days} дней")
        if self.keywords:
            print(f"Активен фильтр по словам: {', '.join(self.keywords)}")
        print(f"{'='*60}\n")

        start_date = datetime.now() - timedelta(days=days)
        print(f"Парсинг канала: @{self.channel_username}")
        print(f"Парсинг сообщений с {start_date.strftime('%d.%m.%Y %H:%M:%S')}")
        print(f"URL: {self.base_url}\n")

        message_count = 0
        all_messages = []

        try:
            html = await self.fetch_page(self.base_url)
            if not html:
                print("Ошибка: Не удалось получить страницу канала")
                print("Убедитесь, что канал публичный и имя указано правильно")
                return

            messages = await self.parse_messages_from_page(html, start_date)
            all_messages.extend(messages)

            print(f"Найдено сообщений на первой странице: {len(messages)}")

            if messages:
                oldest_id = min(msg['id'] for msg in messages)
                empty_pages = 0

                for i in range(50):
                    url = f"{self.base_url}?before={oldest_id}"
                    print(f"\nЗагрузка страницы {i+2}, URL: {url}")
                    html = await self.fetch_page(url)

                    if not html:
                        print(f"Не удалось загрузить страницу {i+2}")
                        break

                    new_messages = await self.parse_messages_from_page(html, start_date)

                    if not new_messages:
                        empty_pages += 1
                        print(f"Пустая страница {i+2} (пустых подряд: {empty_pages})")

                        if empty_pages >= 3:
                            print(f"Остановка после {empty_pages} пустых страниц подряд")
                            break

                        soup = BeautifulSoup(html, 'html.parser')
                        msgs = soup.find_all('div', class_='tgme_widget_message')
                        if msgs:
                            last_msg = msgs[-1]
                            link = last_msg.get('data-post', '')
                            if '/' in link:
                                oldest_id = int(link.split('/')[-1])
                                continue
                        break
                    else:
                        empty_pages = 0

                    all_messages.extend(new_messages)
                    oldest_id = min(msg['id'] for msg in new_messages)

                    print(f"Загружено сообщений: {len(all_messages)} (страница {i+2}, новых: {len(new_messages)})")

                    if len(all_messages) >= MAX_MESSAGES:
                        print(f"Достигнут лимит в {MAX_MESSAGES} сообщений")
                        break

                    await asyncio.sleep(2)

            all_messages.sort(key=lambda x: x['datetime'])

            for msg in all_messages:
                if msg['datetime'] < start_date:
                    continue

                row = [msg['id'], msg['date'], msg['time'], msg['author'],
                       msg['text'], msg['views'], msg['link']]
                self.worksheet.append(row)
                message_count += 1

                print(f"Добавлено сообщение ID: {msg['id']} от {msg['date']} {msg['time']}")

                if message_count % 50 == 0:
                    self.save_excel()

            self.save_excel()
            print(f"\n{'='*60}")
            print(f"Парсинг завершен! Всего обработано: {message_count} сообщений")
            if self.keywords:
                print(f"Фильтр по словам [{', '.join(self.keywords)}] применен")
                print(f"Сообщений соответствует фильтру: {message_count}")
            print(f"{'='*60}\n")

        except Exception as e:
            print(f"Ошибка при парсинге: {e}")
            import traceback
            traceback.print_exc()

    async def run(self):
        try:
            print("Инициализация парсера...")
            self.session = aiohttp.ClientSession()
            self.init_excel()
            await self.parse_historical_messages(days=PARSE_PERIOD_DAYS)
        except KeyboardInterrupt:
            print("\n\nОстановка парсера...")
        except Exception as e:
            print(f"Критическая ошибка: {e}")
            import traceback
            traceback.print_exc()
        finally:
            if self.workbook:
                self.save_excel()
            if self.session:
                await self.session.close()
            print("Парсер остановлен.")


async def main():
    print("="*60)
    print("Telegram Channel Parser Bot")
    print("="*60)

    if CHANNEL_USERNAME == '@your_channel':
        print("ОШИБКА: Необходимо указать имя канала для парсинга!")
        print("Откройте файл и замените '@your_channel' на имя канала")
        return

    print(f"Канал для парсинга: {CHANNEL_USERNAME}")
    print(f"Период: {PARSE_PERIOD_DAYS} дней")
    print(f"Выходной файл: {EXCEL_FILE}")
    print(f"Максимум сообщений: {MAX_MESSAGES}")
    if FILTER_KEYWORDS:
        print(f"Фильтр по словам: {FILTER_KEYWORDS}")
    print()

    parser = TelegramChannelParser(
        channel_username=CHANNEL_USERNAME,
        excel_file=EXCEL_FILE,
        filter_keywords=FILTER_KEYWORDS
    )

    await parser.run()


if __name__ == '__main__':
    asyncio.run(main())
