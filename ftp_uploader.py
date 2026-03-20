import os
import re
import threading
from ftplib import FTP
from pathlib import Path
import tkinter as tk
import ttkbootstrap as ttk
from tkinter import filedialog, messagebox
from ttkbootstrap.scrolled import ScrolledText
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
        # Общий контейнер
        main = ttk.Frame(self.root)
        main.pack(fill="both", expand=True)

        # ================= FTP =================
        ftp_frame = ttk.LabelFrame(main, text="FTP настройки")
        ftp_frame.grid(row=0, column=0, sticky="ew", pady=5)

        ttk.Label(ftp_frame, text="FTP адрес").grid(row=0, column=0, sticky="w")
        self.ftp_host = ttk.Entry(ftp_frame)
        self.ftp_host.grid(row=0, column=1, sticky="ew", padx=5)

        ttk.Label(ftp_frame, text="Логин").grid(row=1, column=0, sticky="w")
        self.ftp_login = ttk.Entry(ftp_frame)
        self.ftp_login.grid(row=1, column=1, sticky="ew", padx=5)

        ttk.Label(ftp_frame, text="Пароль").grid(row=2, column=0, sticky="w")
        self.ftp_password = ttk.Entry(ftp_frame, show="*")
        self.ftp_password.grid(row=2, column=1, sticky="ew", padx=5)

        ttk.Label(ftp_frame, text="Папка на FTP").grid(row=3, column=0, sticky="w")
        self.base_dir = ttk.Entry(ftp_frame)
        self.base_dir.grid(row=3, column=1, sticky="ew", padx=5)

        ftp_frame.columnconfigure(1, weight=1)

        # ================= ФАЙЛЫ =================
        file_frame = ttk.LabelFrame(main, text="Файлы")
        file_frame.grid(row=1, column=0, sticky="ew", pady=5)

        ttk.Label(file_frame, text="Локальная папка").grid(row=0, column=0, sticky="w")
        self.local_dir = ttk.Entry(file_frame)
        self.local_dir.grid(row=0, column=1, sticky="ew", padx=5)

        ttk.Button(file_frame, text="Выбрать", command=self.select_local_folder)\
            .grid(row=0, column=2, padx=5)

        ttk.Label(file_frame, text="Код района").grid(row=1, column=0, sticky="w")
        self.district_code = ttk.Entry(file_frame, width=10)
        self.district_code.grid(row=1, column=1, sticky="w", padx=5)

        ttk.Label(file_frame, text="Доп. путь").grid(row=2, column=0, sticky="w")
        self.inner_path = ttk.Entry(file_frame)
        self.inner_path.grid(row=2, column=1, sticky="ew", padx=5)

        self.progress = ttk.Progressbar(main, mode='determinate')
        self.progress.grid(row=5, column=0, sticky="ew", padx=5, pady=5)

        file_frame.columnconfigure(1, weight=1)

        # ================= ПЕРЕИМЕНОВАНИЕ =================
        rename_frame = ttk.LabelFrame(main, text="Переименование")
        rename_frame.grid(row=2, column=0, sticky="ew", pady=5)

        ttk.Label(rename_frame, text="Маска файла").grid(row=0, column=0, sticky="w")

        self.filename_mask = ttk.Entry(rename_frame)
        self.filename_mask.grid(row=0, column=1, sticky="ew", padx=5)

        ttk.Button(rename_frame, text="?", width=3,
                command=self.show_mask_help).grid(row=0, column=2)

        rename_frame.columnconfigure(1, weight=1)

        # ================= КНОПКИ =================
        btn_frame = ttk.Frame(main)
        btn_frame.grid(row=3, column=0, pady=10)

        ttk.Button(btn_frame, text="Предпросмотр",
                command=self.preview_rename).grid(row=0, column=0, padx=5)

        ttk.Button(btn_frame, text="Загрузить",
                command=self.start_upload).grid(row=0, column=1, padx=5)

        # ================= ЛОГ =================
        log_frame = ttk.LabelFrame(main, text="Лог")
        log_frame.grid(row=4, column=0, sticky="nsew", pady=5, padx=5)

        self.log_area = ScrolledText(log_frame, height=12)
        self.log_area.pack(fill="both", expand=True)

        main.columnconfigure(0, weight=1)
        main.rowconfigure(4, weight=1)

        self.status_label = ttk.Label(main, text="Готов")
        self.status_label.grid(row=6, column=0, sticky="w", padx=5)

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


    def validate_inputs(self):
        errors = []

        if not self.ftp_host.get().strip():
            errors.append("FTP адрес")

        if not self.ftp_login.get().strip():
            errors.append("Логин")

        if not self.local_dir.get().strip():
            errors.append("Локальная папка")

        if errors:
            messagebox.showerror("Ошибка", "Заполните поля:\n" + "\n".join(errors))
            return False

        return True


    def log(self, message, level="info"):
        self.root.after(0, lambda: self._append_log(message, level))

    def _append_log(self, message, level="info"):
        self.log_area.insert(tk.END, message + "\n", level)

        self.log_area.tag_config("error", foreground="red")
        self.log_area.tag_config("success", foreground="green")
        self.log_area.tag_config("info", foreground="black")

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
            self.log("Локальная папка не найдена.", level="error")
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

            self.log(f"Подключение к FTP серверу {host} успешно.", level="success")
        except Exception as e:
            self.log(f"Ошибка подключения: {e}", level="error")
            return

        try:

            if base_dir:
                ftp.cwd(base_dir)
                self.log(f"Переход в папку {base_dir}", level="info")

            working_dir = ftp.pwd()  # 🔹 абсолютный путь

            all_dirs = []
            ftp.retrlines("NLST", lambda line: all_dirs.append(line.strip()))

        except Exception as e:
            self.log(f"Ошибка перехода в директорию: {e}", level="error")
            ftp.quit()
            return
        
        files = [f for f in os.listdir(local_dir) if os.path.isfile(os.path.join(local_dir, f))]

        total = len(files)

        self.root.after(0, lambda: self.progress.configure(maximum=total, value=0))

        for i, file in enumerate(files, start=1):

            full_path = os.path.join(local_dir, file)
            if not os.path.isfile(full_path):
                continue

            ou_number = self.extract_ou_number(file)
            special_folder = None

            if not ou_number:
                special_folder = self.extract_special_school(file)

                if not special_folder:
                    self.log(f"Пропущен '{file}' — учреждение не распознано.", level="error")
                    continue

            remote_folder = None

            # если это специальная школа (не число)
            if ou_number and not str(ou_number).isdigit():

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
                    self.log(f"Папка для ОУ {ou_number} не найдена.", level="error")
                else:
                    self.log(f"Папка для учреждения '{special_folder}' не найдена.", level="error")
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
                self.log(f"Ошибка загрузки '{file}': {e}", level="error")
            
            self.root.after(0, lambda i=i: self.progress.configure(value=i))

        ftp.quit()
        self.log("Загрузка завершена.", level="success")

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
            self.log("Локальная папка не найдена.", level="error")
            return

        if not filename_mask:
            self.log("Маска имени файла не указана.", level="info")
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
            self.log("Нет файлов для предварительного просмотра.", level="info")
            return

        self.show_preview_window(preview_data, local_dir)

    def show_preview_window(self, preview_data, local_dir):
        preview_window = ttk.Toplevel(self.root)
        preview_window.title("Предварительный просмотр переименования")
        preview_window.geometry("800x500")

        tree = ttk.Treeview(preview_window, columns=("old", "new"), show="headings")

        tree.heading("old", text="Старое имя")
        tree.heading("new", text="Новое имя")

        tree.column("old", width=350)
        tree.column("new", width=350)

        for old, new in preview_data:
            tree.insert("", "end", values=(old, new))

        tree.pack(fill="both", expand=True)

        def confirm():
            self.execute_rename(preview_data, local_dir)
            preview_window.destroy()

        ttk.Button(preview_window, text="Подтвердить переименование", command=confirm, bootstyle="success").pack(pady=10)

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
                self.log(f"Ошибка переименования '{old}': {e}", level="error")

        self.log("Переименование завершено.", level="success")

    def show_mask_help(self):
        help_window = ttk.Toplevel(self.root)
        help_window.title("Подсказка по маске имени файла")
        help_window.geometry("600x400")

        text = ScrolledText(help_window, wrap="word")
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
        text.config()

    def set_ui_state(self, state):
        for widget in self.root.winfo_children():
            try:
                widget.configure(state=state)
            except:
                pass    

    def start_upload(self):
        if not self.validate_inputs():
            return

        self.set_ui_state("disabled")
        self.status_label.config(text="Загрузка...")
        threading.Thread(target=self._upload_wrapper, daemon=True).start()

    def _upload_wrapper(self):
        
        try:
            self.upload_files()
        finally:
            self.root.after(0, lambda: self.set_ui_state("normal"))
            self.root.after(0, lambda: self.status_label.config(text="Готов"))


if __name__ == "__main__":
    root = ttk.Window(themename="cosmo")
    app = FTPUploader(root)
    root.mainloop()