import asyncio
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import aiosmtplib
import pandas as pd
from dotenv import load_dotenv

load_dotenv()

# Путь к загруженному файлу
file_path = 'content/file.xlsx'

# Чтение данных
data = pd.read_excel(file_path)
print(data.head())

# Группировка данных
grouped_data = data.groupby('Категории')['Tекст обращения'].apply(list).reset_index()
print(grouped_data)


async def send_email(sender_address, sender_pass, receiver_address, subject, content):
    message = MIMEMultipart()
    message["From"] = sender_address
    message["To"] = receiver_address
    message["Subject"] = subject
    message.attach(MIMEText(content, "plain"))

    # Асинхронное создание сессии SMTP и отправка сообщения
    async with aiosmtplib.SMTP(hostname="smtp.mail.ru", port=465, use_tls=True) as session:
        await session.login(sender_address, sender_pass)
        await session.send_message(message)
        print(f'Письмо с темой "{subject}" отправлено')


async def main():
    sender_address = os.getenv('EMAIL_USER')  # Замените на ваш адрес
    sender_pass = os.getenv('EMAIL_PASS')  # Замените на ваш пароль
    receiver_address = os.getenv('EMAIL_RECEIVER')  # Замените на адрес получателя

    tasks = []
    for category, appeals in grouped_data.values:
        subject = f'Тема письма {category}'
        content = "\n".join([f"Обращение {i + 1}: {appeal}" for i, appeal in enumerate(appeals)])
        tasks.append(send_email(sender_address, sender_pass, receiver_address, subject, content))

    await asyncio.gather(*tasks)


# Запуск асинхронной функции
asyncio.run(main())
