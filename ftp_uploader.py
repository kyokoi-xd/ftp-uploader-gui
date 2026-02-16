import os
import re
import threading
from ftplib import FTP
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

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
        host = self.ftp_host.get()
        login = self.ftp_login.get()
        password = self.ftp_password.get()
        base_dir = self.base_dir.get()
        local_dir = self.local_dir.get()
        filename_mask = self.filename_mask.get()

        if not os.path.isdir(local_dir):
            self.log("Локальная папка не найдена.")
            return

        try:
            ftp = FTP(host)
            ftp.login(login, password)
            self.log(f"Подключение к FTP серверу {host} успешно.")
        except Exception as e:
            self.log(f"Ошибка подключения к FTP серверу: {e}")
            return

        # Получаем список папок на сервере один раз
        try:
            ftp.cwd(base_dir)
            all_dirs = ftp.nlst()
        except Exception as e:
            self.log(f"Ошибка получения списка папок на FTP: {e}")
            ftp.quit()
            return

        for file in os.listdir(local_dir):

            # Только Excel
            if not file.lower().endswith((".xlsx", ".xls")):
                continue

            full_path = os.path.join(local_dir, file)
            if not os.path.isfile(full_path):
                continue

            # Извлекаем номер ОУ из имени файла
            ou_number = self.extract_ou_number(file)
            if not ou_number:
                self.log(f"Пропущен файл '{file}' - не найден номер ОУ.")
                continue

            # Ищем папку на сервере
            remote_dir = None
            pattern = rf'(?<!\d){ou_number}(?!\d)'  # число не должно быть частью другого числа
            for d in all_dirs:
                if re.search(pattern, d):
                    remote_dir = f"{base_dir}/{d}"
                    break

            if not remote_dir:
                self.log(f"На FTP не найдена папка для ОУ {ou_number}")
                continue

            # Формируем имя файла по маске, сохраняем расширение
            ext = Path(file).suffix
            if not filename_mask.strip():
                new_filename = file
            else:
                # Добавляем расширение отдельно, если не указан {ext} в маске
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
                # Загружаем файл напрямую в папку на сервере
                remote_file_path = f"{remote_dir}/{new_filename}"
                with open(full_path, 'rb') as f:
                    ftp.storbinary(f'STOR {remote_file_path}', f)

                self.log(f"{file} → {remote_file_path}")

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