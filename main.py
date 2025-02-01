import pyautogui
import pygetwindow as gw
import time
import re
import subprocess
import pyperclip
import imaplib
import email
import quopri
from bs4 import BeautifulSoup
import sys
import os


# Функция для получения всех данных из файла accounts.txt
def get_all_accounts(file_name):
    with open(file_name, 'r', encoding='utf-8') as file:
        lines = file.readlines()

    accounts = []
    for line in lines:
        account_data = line.strip()
        # Пытаемся разбить на ровно 5 частей
        fields = account_data.split(":", 4)
        if len(fields) == 5:
            steam_login, steam_password, email_login, email_password, profile_link = fields
            accounts.append((steam_login, steam_password, email_login, email_password, profile_link))
        else:
            print(f"Неправильный формат строки: {account_data}")

    return accounts if accounts else None


# Функция для поиска пятизначного кода в тексте письма
def find_code_in_message(message):
    codes = re.findall(r'\b[A-Z0-9]{5}\b', message)
    for code in codes:
        if code != "98009":  # Пропускаем код "98009"
            return code
    return None


# Функция для декодирования и обработки Quoted-Printable
def decode_email_part(part):
    if part.get('Content-Transfer-Encoding') == 'quoted-printable':
        decoded_part = quopri.decodestring(part.get_payload()).decode(part.get_content_charset() or 'utf-8')
    else:
        decoded_part = part.get_payload(decode=True).decode(part.get_content_charset() or 'utf-8')
    return decoded_part


# Функция для получения кода из почты
def get_email_code(email_login, email_password):
    imap_server = "imap.firstmail.ltd"
    port = 993

    mail = imaplib.IMAP4_SSL(imap_server, port)

    try:
        # Авторизуемся на почте
        mail.login(email_login, email_password)
        mail.select("INBOX")

        status, messages = mail.search(None, "ALL")
        if status != "OK":
            print("Не удалось получить список писем.")
            return None

        messages = messages[0].split()
        num_messages = len(messages)

        # Проверяем последние три письма
        for i in range(1, 4):
            if num_messages >= i:
                status, msg = mail.fetch(messages[-i], "(RFC822)")
                if status != "OK":
                    print(f"Не удалось извлечь письмо №{i}.")
                    continue

                raw_message = msg[0][1]
                email_message = email.message_from_bytes(raw_message)

                if email_message.is_multipart():
                    for part in email_message.walk():
                        content_type = part.get_content_type()
                        content_disposition = str(part.get("Content-Disposition"))

                        if "text/html" in content_type and "attachment" not in content_disposition:
                            body = decode_email_part(part)
                            soup = BeautifulSoup(body, 'html.parser')
                            text = soup.get_text()
                            code_match = find_code_in_message(text)
                            if code_match:
                                return code_match
                else:
                    body = decode_email_part(email_message)
                    code_match = find_code_in_message(body)
                    if code_match:
                        return code_match

        print("Код не найден в последних 3 письмах.")
    finally:
        mail.logout()

    return None


# Основной скрипт для работы с Steam Desktop Authenticator
def process_sda(steam_login, steam_password, email_login, email_password, profile_link):
    # Определяем путь к текущей директории
    if getattr(sys, 'frozen', False):  # Если скомпилирован в .exe
        script_dir = os.path.dirname(sys.executable)
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))

    # Путь к Steam Desktop Authenticator.exe, лежащему в папке sda
    sda_dir = os.path.join(script_dir, "sda")
    sda_path = os.path.join(sda_dir, "Steam Desktop Authenticator.exe")

    # Проверяем, существует ли файл SDA
    if not os.path.exists(sda_path):
        print(f"Файл Steam Desktop Authenticator.exe не найден в директории {sda_dir}.")
    else:
        print(f"Файл Steam Desktop Authenticator.exe найден, путь: {sda_path}.")

    # Ищем окно SDA
    window = gw.getWindowsWithTitle("Steam Desktop Authenticator")
    if not window:
        subprocess.Popen(sda_path)
        time.sleep(3)
    else:
        sda_window = window[0]
        sda_window.activate()
        time.sleep(2)

    # Координаты элементов интерфейса SDA
    ok_on_start_button = (1123, 611)
    SetupNewAccount_button = (900, 350)
    login_field = (900, 476)
    password_field = (900, 509)
    login_button = (1050, 590)
    ok_first_button = (1030, 600)
    ok_second_button = (1110, 600)
    input_code_field = (800, 600)
    enter_email_code_button = (800, 630)
    input_r_code_field = (800, 600)
    enter_r_code_button = (800, 630)
    ok_r_code_button = (1110, 600)

    # Начинаем добавление нового аккаунта в SDA
    pyautogui.click(ok_on_start_button)
    time.sleep(1)
    pyautogui.click(SetupNewAccount_button)
    time.sleep(1)
    pyautogui.click(login_field)
    pyautogui.typewrite(steam_login)
    time.sleep(1)
    pyautogui.click(password_field)
    pyautogui.typewrite(steam_password)
    time.sleep(1)
    pyautogui.click(login_button)
    time.sleep(3)
    pyautogui.click(ok_first_button)
    time.sleep(3)

    # Получаем код из почты
    email_code = get_email_code(email_login, email_password)
    if not email_code:
        print("Не удалось получить код из почты.")
        return

    # Вводим код из почты
    pyautogui.click(input_code_field)
    time.sleep(3)
    pyautogui.typewrite(email_code)
    time.sleep(1)
    pyautogui.click(enter_email_code_button)
    time.sleep(1)

    # Продолжаем
    pyautogui.click(enter_r_code_button)
    time.sleep(2)
    pyautogui.click(ok_first_button)
    time.sleep(2)

    # Копируем текст из окна и ищем R-код
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(0.5)
    window_text = pyperclip.paste()

    match = re.search(r'R\d+', window_text)
    if not match:
        print("R-код не найден.")
        return
    revocation_code = match.group()
    print(f"Найден R-код: {revocation_code}")

    time.sleep(6)  # Ждём, пока придёт код на почту

    pyautogui.click(ok_second_button)  # Нажимаем кнопку "OK"
    time.sleep(10)

    # Получаем код из почты повторно
    email_code = get_email_code(email_login, email_password)
    if not email_code:
        print("Не удалось получить код из почты.")
        return

    # Вводим второй код
    pyautogui.click(input_code_field)
    time.sleep(3)
    pyautogui.typewrite(email_code)
    time.sleep(1)
    pyautogui.click(enter_email_code_button)
    time.sleep(1)

    # Вводим R-код
    pyautogui.click(input_r_code_field)
    pyautogui.typewrite(revocation_code)
    time.sleep(1)
    pyautogui.click(enter_r_code_button)
    time.sleep(2.5)
    pyautogui.click(ok_r_code_button)
    time.sleep(1)

    # Сохраняем данные в файл ready_accounts.txt
    with open("ready_accounts.txt", "a", encoding='utf-8') as file:
        file.write(f"{steam_login}:{steam_password}:{email_login}:{email_password}:{profile_link}:{revocation_code}\n")
        print("Аккаунт добавлен в SDA и записан в ready_accounts.txt")


# Запускаем основной цикл
if __name__ == "__main__":
    accounts = get_all_accounts("accounts.txt")
    if accounts:
        for account_data in accounts:
            steam_login, steam_password, email_login, email_password, profile_link = account_data
            process_sda(steam_login, steam_password, email_login, email_password, profile_link)