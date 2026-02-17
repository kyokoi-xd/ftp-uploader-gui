import os
import re
import threading
from ftplib import FTP
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from dotenv import load_dotenv

load_dotenv()

VALID_OU_NUMBERS = {
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
    "333"
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
        
        tk.Label(text="Локальная папка:").pack(anchor='w')
        self.local_dir = tk.Entry(width=60)
        self.local_dir.pack()
        tk.Button(text="Выбрать папку", command=self.select_local_folder).pack(pady=5)

        tk.Label(text="Маска имени файла").pack(anchor='w')
        self.filename_mask = tk.Entry(width=60)
        self.filename_mask.pack()

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
    
    def find_remote_folder(self, ftp, base_dir, ou_number):
        try:
            ftp.cwd(base_dir)
            items = ftp.nlst()

            pattern = rf'(?<!\d){ou_number}(?!\d)'

            for item in items:
                if re.search(pattern, item):
                    return f"{base_dir}/{item}"

            return None

        except Exception as e:
            self.log(f"Ошибка поиска папки на FTP: {e}")
            return None


 
    def upload_files(self):
        host = self.ftp_host.get().strip()
        login = self.ftp_login.get().strip()
        password = self.ftp_password.get().strip()
        base_dir = self.base_dir.get().strip()
        local_dir = self.local_dir.get().strip()
        filename_mask = self.filename_mask.get().strip()

        if not os.path.isdir(local_dir):
            self.log("Локальная папка не найдена.")
            return

        try:
            ftp = FTP(host)
            ftp.login(login, password)
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

            all_dirs = ftp.nlst()

        except Exception as e:
            self.log(f"Ошибка перехода в директорию: {e}")
            ftp.quit()
            return

        for file in os.listdir(local_dir):

            if not file.lower().endswith((".xlsx", ".xls")):
                continue

            full_path = os.path.join(local_dir, file)
            if not os.path.isfile(full_path):
                continue

            ou_number = self.extract_ou_number(file)
            if not ou_number:
                self.log(f"Пропущен '{file}' — номер ОУ не найден.")
                continue

            remote_folder = None
            pattern = rf'(?<!\d){ou_number}(?!\d)'

            for d in all_dirs:
                if re.search(pattern, d):
                    remote_folder = d
                    break

            if not remote_folder:
                self.log(f"Папка для ОУ {ou_number} не найдена.")
                continue

            ext = Path(file).suffix

            if not filename_mask:
                new_filename = file
            else:
                if '{ext}' not in filename_mask:
                    new_filename = filename_mask.format(
                        ou=ou_number,
                        original=file
                    ) + ext
                else:
                    new_filename = filename_mask.format(
                        ou=ou_number,
                        original=file,
                        ext=ext
                    )

            try:
                # 🔹 всегда используем абсолютный путь
                ftp.cwd(f"{working_dir}/{remote_folder}")

                with open(full_path, 'rb') as f:
                    ftp.storbinary(f"STOR {new_filename}", f)

                self.log(f"{file} → {remote_folder}/{new_filename}")

                # 🔹 возвращаемся обратно в working_dir
                ftp.cwd(working_dir)

            except Exception as e:
                self.log(f"Ошибка загрузки '{file}': {e}")

        ftp.quit()
        self.log("Загрузка завершена.")






    
    def start_upload(self):
        threading.Thread(target=self.upload_files).start()


if __name__ == "__main__":
    root = tk.Tk()
    app = FTPUploader(root)
    root.mainloop()