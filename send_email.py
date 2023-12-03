import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from dotenv import load_dotenv
import os
import pandas as pd

load_dotenv()
# Путь к загруженному файлу
file_path = 'content/file.xlsx'

data = pd.read_excel(file_path)

print(data.head())

grouped_data = data.groupby('Категории')['Tекст обращения'].apply(list).reset_index()

print(grouped_data)

# Установите параметры
sender_address = os.getenv('EMAIL_USER')
sender_pass = os.getenv('EMAIL_PASS')
receiver_address = os.getenv('EMAIL_RECEIVER')

# Создайте сессию SMTP и отправьте сообщение
session = smtplib.SMTP('smtp.mail.ru', 587)  # Используйте подходящий SMTP сервер
session.starttls()  # Включите безопасный транспортный слой
session.login(sender_address, sender_pass)  # Авторизуйтесь

for category, appeals in grouped_data.values:
    # Создайте новый экземпляр сообщения для каждого письма
    message = MIMEMultipart()
    message['From'] = sender_address
    message['To'] = receiver_address
    message['Subject'] = f'Тема письма {category}'

    # Тело письма
    mail_content = "\n".join([f"Обращение {i + 1}: {appeal}" for i, appeal in enumerate(appeals)])
    message.attach(MIMEText(mail_content, 'plain'))

    # Отправка сообщения
    text = message.as_string()
    session.sendmail(sender_address, receiver_address, text)
    print(f'Письмо {category} отправлено')

session.quit()

print('Почта успешно отправлена')
