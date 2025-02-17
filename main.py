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
import psutil
import sys
import os

def is_process_running(process_name):
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] == process_name:
            return True
    return False

def kill_process_by_name(process_name):
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] == process_name:
            proc.terminate()

def restart_sda(sda_path):
    print("[INFO] Перезапуск Steam Desktop Authenticator...")
    if is_process_running("Steam Desktop Authenticator.exe"):
        kill_process_by_name("Steam Desktop Authenticator.exe")
    time.sleep(5)
    subprocess.Popen(sda_path)
    time.sleep(5)

def get_all_accounts(file_name):
    with open(file_name, 'r', encoding='utf-8') as file:
        lines = file.readlines()

    accounts = []
    for line in lines:
        account_data = line.strip()
        if not account_data:
            continue
        # Разбиваем строку, минимум 4 поля (логин/пароль стима и логин/пароль почты)
        fields = account_data.split(":")
        if len(fields) < 4:
            print(f"Неправильный формат строки (не хватает полей): {account_data}")
            continue

        steam_login, steam_password, email_login, email_password = fields[:4]
        # Пятая часть (profile_link) — опциональна
        profile_link = fields[4] if len(fields) >= 5 else None

        accounts.append((steam_login, steam_password, email_login, email_password, profile_link))

    return accounts if accounts else None

def remove_account_line(steam_login, steam_password, email_login, email_password, profile_link, file_name="accounts.txt"):
    fields_to_remove = [steam_login, steam_password, email_login, email_password]
    if profile_link:
        fields_to_remove.append(profile_link)
    line_to_remove = ":".join(fields_to_remove) + "\n"

    # Читаем все строки
    with open(file_name, "r", encoding="utf-8") as f:
        lines = f.readlines()

    # Перезаписываем, исключая нужную строку
    with open(file_name, "w", encoding="utf-8") as f:
        for line in lines:
            if line.strip() != line_to_remove.strip():
                f.write(line)

def find_code_in_message(message):
    codes = re.findall(r'\b[A-Z0-9]{5}\b', message)
    for code in codes:
        # Пропустим код "98009"
        if code != "98009":
            return code
    return None

def decode_email_part(part):
    if part.get('Content-Transfer-Encoding') == 'quoted-printable':
        decoded_part = quopri.decodestring(part.get_payload()).decode(part.get_content_charset() or 'utf-8')
    else:
        decoded_part = part.get_payload(decode=True).decode(part.get_content_charset() or 'utf-8')
    return decoded_part

def get_email_code(email_login, email_password):
    imap_server = "imap.firstmail.ltd"
    port = 993

    mail = imaplib.IMAP4_SSL(imap_server, port)

    try:
        mail.login(email_login, email_password)
        mail.select("INBOX")

        status, messages = mail.search(None, "ALL")
        if status != "OK":
            print("Не удалось получить список писем.")
            return None

        messages = messages[0].split()
        num_messages = len(messages)

        # Проверяем последние три письма с конца
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
                            code_match = find_code_in_message(BeautifulSoup(body, 'html.parser').get_text())
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

def process_sda(steam_login, steam_password, email_login, email_password, profile_link):
    # Определяем путь к текущей директории
    if getattr(sys, 'frozen', False):  # Если скомпилирован в .exe
        script_dir = os.path.dirname(sys.executable)
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))

    # Путь к Steam Desktop Authenticator.exe (в папке "sda")
    sda_dir = os.path.join(script_dir, "sda")
    sda_path = os.path.join(sda_dir, "Steam Desktop Authenticator.exe")

    if not os.path.exists(sda_path):
        print(f"Файл Steam Desktop Authenticator.exe не найден по пути: {sda_path}")
        return
    else:
        print(f"SDA найден: {sda_path}")

    # Пытаемся найти уже открытое окно SDA
    window = gw.getWindowsWithTitle("Steam Desktop Authenticator")
    if not window:
        subprocess.Popen(sda_path)
        time.sleep(3)
    else:
        window[0].activate()
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

    # Шаги добавления аккаунта
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

    # Получаем код из почты (первая попытка)
    email_code = get_email_code(email_login, email_password)
    if not email_code:
        print("Не удалось получить код из почты. Перезапускаем SDA и пытаемся снова...")
        restart_sda(sda_path)
        return process_sda(steam_login, steam_password, email_login, email_password, profile_link)

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
        print("R-код не найден. Перезапускаем SDA и пытаемся снова...")
        restart_sda(sda_path)
        return process_sda(steam_login, steam_password, email_login, email_password, profile_link)

    revocation_code = match.group()
    print(f"Найден R-код: {revocation_code}")

    time.sleep(6)  # Ждём, пока придёт второй код на почту
    pyautogui.click(ok_second_button)
    time.sleep(10)

    # Снова получаем код из почты (вторая попытка)
    email_code = get_email_code(email_login, email_password)
    if not email_code:
        print("Не удалось получить второй код из почты. Перезапускаем SDA и пытаемся снова...")
        restart_sda(sda_path)
        return process_sda(steam_login, steam_password, email_login, email_password, profile_link)

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

    # Запишем данные в ready_accounts.txt
    with open("ready_accounts.txt", "a", encoding='utf-8') as file:
        # Собираем поля гибко (пятый — профиль, если есть)
        fields_to_write = [steam_login, steam_password, email_login, email_password]
        if profile_link:
            fields_to_write.append(profile_link)
        fields_to_write.append(revocation_code)
        file.write(":".join(fields_to_write) + "\n")

    print("Аккаунт добавлен в SDA и записан в ready_accounts.txt")

    # Удаляем использованную строку из accounts.txt
    remove_account_line(steam_login, steam_password, email_login, email_password, profile_link, "accounts.txt")

if __name__ == "__main__":
    accounts = get_all_accounts("accounts.txt")
    if accounts:
        for account_data in accounts:
            steam_login, steam_password, email_login, email_password, profile_link = account_data
            process_sda(steam_login, steam_password, email_login, email_password, profile_link)
    else:
        print("[INFO] Не найдены аккаунты в файле accounts.txt")