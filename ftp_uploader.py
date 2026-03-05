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

SPECIAL_SCHOOLS = {
    "clschool39": ["шря", "clschool39"],
    "deltaschool": ["дельта", "deltaschool"],
    "SVU": ["сву", "svu"]
}
SPECIAL_NUMBERS = {
    "5661": "SVU",
    "5663": "clschool39",
    "5664": "deltaschool"
}
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

        tk.Label(text="Код района (если есть):").pack(anchor='w')
        self.district_code = tk.Entry(width=10)
        self.district_code.pack(anchor='w')

        tk.Label(text="Доп. путь внутри папки школы:").pack(anchor='w')
        self.inner_path = tk.Entry(width=60)
        self.inner_path.pack()
        
        tk.Label(text="Локальная папка:").pack(anchor='w')
        self.local_dir = tk.Entry(width=60)
        self.local_dir.pack()
        tk.Button(text="Выбрать папку", command=self.select_local_folder).pack(pady=5)

        mask_frame = tk.Frame(self.root)
        mask_frame.pack(anchor='w')

        tk.Label(mask_frame, text="Маска имени файла").pack(side="left")

        tk.Button(mask_frame,text=" ? ",command=self.show_mask_help,bg="#444",fg="white").pack(side="left", padx=5)

        self.filename_mask = tk.Entry(self.root, width=60)
        self.filename_mask.pack()

        tk.Button(text="Предварительный просмотр",command=self.preview_rename,bg="blue",fg="white").pack(pady=5)
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

        district_code = self.district_code.get().strip()

        # 1️⃣ Приоритет — номер после №
        match = re.search(r'(?:№|N|N°)\s*(\d+)', text)
        if match:
            number = match.group(1)

            # проверка обычного номера
            if number in VALID_OU_NUMBERS:
                return number

            # проверка номера с кодом района
            if district_code and number.startswith(district_code):
                possible = number[len(district_code):]
                if possible in VALID_OU_NUMBERS:
                    return possible

        # 2️⃣ проверяем все числа
        all_numbers = re.findall(r'\d+', text)

        for number in all_numbers:

            if number in SPECIAL_NUMBERS:
                return SPECIAL_NUMBERS[number]
            
            if number in VALID_OU_NUMBERS:
                return number

            if district_code and number.startswith(district_code):
                possible = number[len(district_code):]

                if possible in VALID_OU_NUMBERS:
                    return possible

        return None
    
 
    def upload_files(self):
        host = self.ftp_host.get().strip()
        login = self.ftp_login.get().strip()
        password = self.ftp_password.get().strip()
        base_dir = self.base_dir.get().strip()
        local_dir = self.local_dir.get().strip()
        inner_path = self.inner_path.get().strip().strip("/")


        if not os.path.isdir(local_dir):
            self.log("Локальная папка не найдена.")
            return

        try:
            ftp = FTP()
            ftp.encoding = "cp1251"   # ОБЯЗАТЕЛЬНО ДО connect/login

            ftp.connect(host, 21, timeout=15)
            ftp.login(login, password)

            # отключаем UTF8 если сервер его рекламирует
            try:
                ftp.sendcmd("OPTS UTF8 OFF")
            except:
                pass

            ftp.voidcmd("TYPE I")

            self.log(f"Подключение к FTP серверу {host} успешно.")
        except Exception as e:
            self.log(f"Ошибка подключения: {e}")
            return

        try:

            if base_dir:
                ftp.cwd(base_dir)
                self.log(f"Переход в папку {base_dir}")

            working_dir = ftp.pwd()  # 🔹 абсолютный путь

            all_dirs = []
            ftp.retrlines("NLST", lambda line: all_dirs.append(line.strip()))

        except Exception as e:
            self.log(f"Ошибка перехода в директорию: {e}")
            ftp.quit()
            return

        for file in os.listdir(local_dir):

            full_path = os.path.join(local_dir, file)
            if not os.path.isfile(full_path):
                continue

            ou_number = self.extract_ou_number(file)
            special_folder = None

            if not ou_number:
                special_folder = self.extract_special_school(file)

                if not special_folder:
                    self.log(f"Пропущен '{file}' — учреждение не распознано.")
                    continue

            remote_folder = None

            # если это специальная школа (не число)
            if not str(ou_number).isdigit():

                for d in all_dirs:
                    if d.lower() == str(ou_number).lower():
                        remote_folder = d
                        break

            # если обычная школа
            else:

                for d in all_dirs:
                    numbers = re.findall(r'\d+', d)

                    for num in numbers:
                        if num == str(ou_number):
                            remote_folder = d
                            break

                    if remote_folder:
                        break

  
            if not remote_folder:
                if ou_number:
                    self.log(f"Папка для ОУ {ou_number} не найдена.")
                else:
                    self.log(f"Папка для учреждения '{special_folder}' не найдена.")
                continue

            try:
                # Переход в папку школы
                ftp.cwd(working_dir)
                # Нормализация пути
                if inner_path:
                    normalized_path = inner_path.replace("\\", "/").strip()

                    try:
                        # если путь начинается с / — считаем его абсолютным
                        if normalized_path.startswith("/"):
                            ftp.cwd(normalized_path)
                        else:
                            # иначе — относительно текущей папки школы
                            ftp.cwd(f"{remote_folder}/{normalized_path}")

                    except Exception as e:
                        raise Exception(f"Путь не существует: {normalized_path}")
                else:
                    ftp.cwd(remote_folder)

                # Проверяем существующие файлы
                new_filename = file

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

    def extract_special_school(self, text):
        text_lower = text.lower()

        for folder, keywords in SPECIAL_SCHOOLS.items():
            for keyword in keywords:
                if keyword.lower() in text_lower:
                    return folder

        return None


    def generate_filename(self, filename_mask, file, ou_number, counter=1):
        ext = Path(file).suffix
        original_name = Path(file).stem

        now = datetime.now()
        date_str = now.strftime("%d_%m_%Y")
        datetime_str = now.strftime("%d_%m_%Y_%H_%M_%S")

        try:
            new_name = filename_mask.format(
                ou=ou_number,
                date=date_str,
                datetime=datetime_str,
                original=original_name,
                ext=ext,
                counter=counter
            )
        except KeyError as e:
            raise Exception(f"Неизвестная переменная в маске: {e}")

        # 🔥 УДАЛЯЕМ запрещённые символы
        new_name = new_name.replace("\n", "").replace("\r", "").strip()

        # Удаляем двойные пробелы
        new_name = re.sub(r"\s+", " ", new_name)

        # Убираем запрещённые для FTP символы
        new_name = re.sub(r'[<>:"/\\|?*]', "_", new_name)

        if not new_name.endswith(ext):
            new_name += ext

        return new_name
    
    def preview_rename(self):
        local_dir = self.local_dir.get().strip()
        filename_mask = self.filename_mask.get().strip()

        if not os.path.isdir(local_dir):
            self.log("Локальная папка не найдена.")
            return

        if not filename_mask:
            self.log("Маска имени файла не указана.")
            return

        preview_data = []

        for file in os.listdir(local_dir):
            full_path = os.path.join(local_dir, file)
            if not os.path.isfile(full_path):
                continue

            ou_number = self.extract_ou_number(file)

            if not ou_number:
                special = self.extract_special_school(file)
                if not special:
                    continue
                ou_number = special

            new_name = self.generate_filename(
                filename_mask,
                file,
                ou_number,
                1
            )

            preview_data.append((file, new_name))

        if not preview_data:
            self.log("Нет файлов для предварительного просмотра.")
            return

        self.show_preview_window(preview_data, local_dir)

    def show_preview_window(self, preview_data, local_dir):
        preview_window = tk.Toplevel(self.root)
        preview_window.title("Предварительный просмотр переименования")
        preview_window.geometry("800x500")

        text = scrolledtext.ScrolledText(preview_window, width=100, height=25)
        text.pack(fill="both", expand=True)

        for old, new in preview_data:
            text.insert(tk.END, f"{old}  →  {new}\n")

        def confirm():
            self.execute_rename(preview_data, local_dir)
            preview_window.destroy()

        tk.Button(
            preview_window,
            text="Подтвердить переименование",
            command=confirm,
            bg="green",
            fg="white"
        ).pack(pady=10)

    def execute_rename(self, preview_data, local_dir):
        for old, new in preview_data:
            old_path = os.path.join(local_dir, old)
            new_path = os.path.join(local_dir, new)

            counter = 1
            base_new = new

            while os.path.exists(new_path):
                name = Path(base_new).stem
                ext = Path(base_new).suffix
                new_filename = f"{name}_{counter}{ext}"
                new_path = os.path.join(local_dir, new_filename)
                counter += 1

            try:
                os.rename(old_path, new_path)
                self.log(f"{old} → {os.path.basename(new_path)}")
            except Exception as e:
                self.log(f"Ошибка переименования '{old}': {e}")

        self.log("Переименование завершено.")

    def show_mask_help(self):
        help_window = tk.Toplevel(self.root)
        help_window.title("Подсказка по маске имени файла")
        help_window.geometry("600x400")

        text = scrolledtext.ScrolledText(help_window, wrap="word")
        text.pack(fill="both", expand=True)

        help_text = """
        ДОСТУПНЫЕ ПЕРЕМЕННЫЕ:

        {ou}        — номер школы
        {date}      — текущая дата (ДД_ММ_ГГГГ)
        {datetime}  — дата и время (ДД_ММ_ГГГГ_ЧЧ_ММ_СС)
        {original}  — исходное имя файла без расширения

        ВАЖНО:
        • Расширение добавляется автоматически
        • Не нужно писать .xlsx или .pdf вручную
        • Переменные обязательно писать в фигурных скобках {}

        ПРИМЕРЫ:

        {ou}_отчет_{date}
        {ou}_{original}_{date}
        {ou}_отчет_{datetime}

        Можно писать обычный текст:
        284_отчет_15_03_2025
        """

        text.insert("1.0", help_text)
        text.config(state="disabled")

    def start_upload(self):
        threading.Thread(target=self.upload_files).start()


if __name__ == "__main__":
    root = tk.Tk()
    app = FTPUploader(root)
    root.mainloop()