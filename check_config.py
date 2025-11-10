"""
Скрипт для проверки конфигурации
Проверяет правильность настроек перед запуском основного бота
"""

import sys

# Попробуем импортировать необходимые библиотеки
print("Проверка установленных библиотек...\n")

try:
    import telethon
    print("✅ telethon установлен (версия: {})".format(telethon.__version__))
except ImportError:
    print("❌ telethon НЕ установлен!")
    print("   Установите: pip install telethon")
    sys.exit(1)

try:
    import openpyxl
    print("✅ openpyxl установлен (версия: {})".format(openpyxl.__version__))
except ImportError:
    print("❌ openpyxl НЕ установлен!")
    print("   Установите: pip install openpyxl")
    sys.exit(1)

print("\n" + "="*60)
print("Проверка конфигурации Parser message.py...\n")

# Импортируем настройки из основного файла
try:
    # Читаем файл напрямую, чтобы не запускать весь скрипт
    with open('Parser message.py', 'r', encoding='utf-8') as f:
        content = f.read()

    # Проверяем наличие плейсхолдеров
    errors = []
    warnings = []

    if 'YOUR_API_ID' in content:
        errors.append("API_ID не настроен (содержит 'YOUR_API_ID')")
    else:
        print("✅ API_ID настроен")

    if 'YOUR_API_HASH' in content:
        errors.append("API_HASH не настроен (содержит 'YOUR_API_HASH')")
    else:
        print("✅ API_HASH настроен")

    if 'YOUR_PHONE_NUMBER' in content:
        errors.append("PHONE не настроен (содержит 'YOUR_PHONE_NUMBER')")
    else:
        print("✅ PHONE настроен")

    if '@your_channel' in content:
        errors.append("CHANNEL_USERNAME не настроен (содержит '@your_channel')")
    else:
        print("✅ CHANNEL_USERNAME настроен")

    print("\n" + "="*60)

    if errors:
        print("\n❌ ОШИБКИ КОНФИГУРАЦИИ:\n")
        for i, error in enumerate(errors, 1):
            print(f"   {i}. {error}")
        print("\nИсправьте ошибки в файле 'Parser message.py' перед запуском!")
        print("Пример настройки смотрите в файле 'config_example.py'")
        sys.exit(1)
    else:
        print("\n✅ ВСЕ НАСТРОЙКИ В ПОРЯДКЕ!")
        print("\nМожно запускать бота:")
        print('   python "Parser message.py"')

    if warnings:
        print("\n⚠️  ПРЕДУПРЕЖДЕНИЯ:\n")
        for i, warning in enumerate(warnings, 1):
            print(f"   {i}. {warning}")

except FileNotFoundError:
    print("❌ Файл 'Parser message.py' не найден!")
    sys.exit(1)
except Exception as e:
    print(f"❌ Ошибка при проверке: {e}")
    sys.exit(1)

print("\n" + "="*60)
print("Проверка завершена!")
print("="*60 + "\n")

