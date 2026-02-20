import os
import re
import threading
from ftplib import FTP
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from dotenv import load_dotenv
from datetime import datetime

load_dotenv()

VALID_OU_NUMBERS = {
    "002",
    "2",
    "162",
    "221",
    "223",
    "244",
    "248",
    "249",
    "250",
    "254",
    "261",
    "264",
    "269",
    "277",
    "282",
    "283",
    "284",
    "377",
    "378",
    "379",
    "381",
    "384",
    "386",
    "387",
    "388",
    "389",
    "392",
    "393",
    "397",
    "480",
    "481",
    "493",
    "501",
    "502",
    "503",
    "504",
    "506",
    "538",
    "539",
    "551",
    "565",
    "585",
    "608",
    "654",
    "658",
}

class FTPUploader:
    def __init__(self, root):
        self.root = root
        self.root.title("FTP Uploader")
        self.root.geometry("700x600")
        
        self.create_widgets()
        self.load_env_config()

    def create_widgets(self):
        # FTP Server Details
        tk.Label(text="FTP адрес:").pack(anchor='w')
        self.ftp_host = tk.Entry(width=60)
        self.ftp_host.pack()

        tk.Label(text="Логин:").pack(anchor='w')
        self.ftp_login = tk.Entry(width=60)
        self.ftp_login.pack()

        tk.Label(text="Пароль:").pack(anchor='w')
        self.ftp_password = tk.Entry(width=60, show="*")
        self.ftp_password.pack()

        tk.Label(text="Папка на FTP:").pack(anchor='w')
        self.base_dir = tk.Entry(width=60)
        self.base_dir.pack()

        tk.Label(text="Доп. путь внутри папки школы (необязательно):").pack(anchor='w')
        self.inner_path = tk.Entry(width=60)
        self.inner_path.pack()
        
        tk.Label(text="Локальная папка:").pack(anchor='w')
        self.local_dir = tk.Entry(width=60)
        self.local_dir.pack()
        tk.Button(text="Выбрать папку", command=self.select_local_folder).pack(pady=5)

        tk.Label(text="Маска имени файла").pack(anchor='w')
        self.filename_mask = tk.Entry(width=60)
        self.filename_mask.pack()

        tk.Label(text="Дата (ДД_ММ_ГГГГ) — можно оставить пустым:").pack(anchor='w')
        self.custom_date = tk.Entry(width=60)
        self.custom_date.pack()


        tk.Button(text="Загрузить", command=self.start_upload, bg='green', fg="white").pack(pady=10)

        tk.Label(text="Лог:").pack(anchor='w')
        self.log_area = scrolledtext.ScrolledText(width=80, height=15)
        self.log_area.pack(fill="both", expand=True)

    def load_env_config(self):
        host = os.getenv("FTP_HOST", "")
        login = os.getenv("FTP_LOGIN", "")
        password = os.getenv("FTP_PASSWORD", "")

        if host:
            self.ftp_host.insert(0, host)

        if login:
            self.ftp_login.insert(0, login)

        if password:
            self.ftp_password.insert(0, password)



    def log(self, message):
        self.root.after(0, lambda: self._append_log(message))

    def _append_log(self, message):
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)


    def select_local_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.local_dir.delete(0, tk.END)
            self.local_dir.insert(0, folder)
    
    def extract_ou_number(self, text):
        # 1️⃣ Приоритет — номер после № / N / N°
        match = re.search(r'(?:№|N|N°)\s*(\d+)', text)
        if match:
            number = match.group(1)
            if number in VALID_OU_NUMBERS:
                return number

        # 2️⃣ Иначе проверяем все числа в строке
        all_numbers = re.findall(r'\d+', text)

        for number in all_numbers:
            if number in VALID_OU_NUMBERS:
                return number

        return None
    
 
    def upload_files(self):
        host = self.ftp_host.get().strip()
        login = self.ftp_login.get().strip()
        password = self.ftp_password.get().strip()
        base_dir = self.base_dir.get().strip()
        local_dir = self.local_dir.get().strip()
        filename_mask = self.filename_mask.get().replace("\n", "").replace("\r", "").strip()
        custom_date = self.custom_date.get().strip()
        inner_path = self.inner_path.get().strip().strip("/")


        if not os.path.isdir(local_dir):
            self.log("Локальная папка не найдена.")
            return

        try:
            ftp = FTP()
            ftp.connect(host, 21, timeout=15)

            ftp.login(login, password)

            # Пытаемся выключить UTF8
            try:
                ftp.sendcmd("OPTS UTF8 OFF")
            except:
                pass

            # Принудительно ставим кодировку
            ftp.encoding = "cp1251"

            ftp.voidcmd("TYPE I")

            self.log(f"Подключение к FTP серверу {host} успешно.")
        except Exception as e:
            self.log(f"Ошибка подключения: {e}")
            return

        try:
            root_dir = ftp.pwd()  # 🔹 запоминаем стартовую папку

            if base_dir:
                ftp.cwd(base_dir)
                self.log(f"Переход в папку {base_dir}")

            working_dir = ftp.pwd()  # 🔹 абсолютный путь

            all_dirs = []

            def collect_dirs(data):
                all_dirs.append(data.decode("cp1251").strip())

            ftp.retrbinary("NLST", collect_dirs)

        except Exception as e:
            self.log(f"Ошибка перехода в директорию: {e}")
            ftp.quit()
            return

        for file in os.listdir(local_dir):

            full_path = os.path.join(local_dir, file)
            if not os.path.isfile(full_path):
                continue

            ou_number = self.extract_ou_number(file)
            if not ou_number:
                self.log(f"Пропущен '{file}' — номер ОУ не найден.")
                continue

            remote_folder = None

            for d in all_dirs:
                numbers = re.findall(r'\d+', d)
                for num in numbers:
                    if int(num) == int(ou_number):
                        remote_folder = d
                        break
                if remote_folder:
                    break

  
            if not remote_folder:
                self.log(f"Папка для ОУ {ou_number} не найдена.")
                continue

            ext = Path(file).suffix
            try:
                # Переход в папку школы
                ftp.cwd(working_dir)
                ftp.cwd(remote_folder)

                # Если указан доп. путь — просто переходи
                if inner_path:
                    parts = inner_path.split("/")

                    for part in parts:
                        dirs = []
                        ftp.retrbinary("NLST", lambda d: dirs.append(d.decode("cp1251").strip()))

                        match = None
                        for d in dirs:
                            if d.strip().lower() == part.strip().lower():
                                match = d
                                break

                        if not match:
                            raise Exception(f"Папка '{part}' не существует на сервере")

                        ftp.cwd(match)

                # Проверяем существующие файлы
                existing_files = ftp.nlst()

                if not filename_mask:
                    new_filename = file
                else:
                    counter = 1
                    new_filename = self.generate_filename(
                        filename_mask,
                        file,
                        ou_number,
                        custom_date,
                        counter
                    )

                    while new_filename in existing_files:
                        counter += 1
                        new_filename = self.generate_filename(
                            filename_mask,
                            file,
                            ou_number,
                            custom_date,
                            counter
                        )


                with open(full_path, 'rb') as f:
                    ftp.storbinary(f"STOR {new_filename}", f)

                full_remote_path = remote_folder
                if inner_path:
                    full_remote_path += "/" + inner_path

                self.log(f"{file} → {full_remote_path}/{new_filename}")

                # Возврат в рабочую директорию
                ftp.cwd(working_dir)

            except Exception as e:
                self.log(f"Ошибка загрузки '{file}': {e}")

        ftp.quit()
        self.log("Загрузка завершена.")


    def generate_filename(self, filename_mask, file, ou_number, custom_date, counter=1):
        ext = Path(file).suffix
        original_name = Path(file).stem

        if custom_date:
            date_str = custom_date
            datetime_str = custom_date
        else:
            now = datetime.now()
            date_str = now.strftime("%d_%m_%Y")
            datetime_str = now.strftime("%d_%m_%Y_%H_%M_%S")

        new_name = filename_mask.format(
            ou=ou_number,
            date=date_str,
            datetime=datetime_str,
            original=original_name,
            ext=ext,
            counter=counter
        )

        # 🔥 УДАЛЯЕМ запрещённые символы
        new_name = new_name.replace("\n", "").replace("\r", "").strip()

        # Удаляем двойные пробелы
        new_name = re.sub(r"\s+", " ", new_name)

        # Убираем запрещённые для FTP символы
        new_name = re.sub(r'[<>:"/\\|?*]', "_", new_name)

        if not new_name.endswith(ext):
            new_name += ext

        return new_name



    def start_upload(self):
        threading.Thread(target=self.upload_files).start()


if __name__ == "__main__":
    root = tk.Tk()
    app = FTPUploader(root)
    root.mainloop()