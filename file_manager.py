import glob
import os.path
from tkinter import filedialog, messagebox, simpledialog

import pandas as pd
from ttkbootstrap.dialogs import Messagebox


class FileManager:
    CONFIG_FILE = "config.txt"
    EMAIL_RECEIVERS = "email_receivers"
    BASE_PATH_KW = "base_path_kw"
    BASE_PATH_MERGE = "base_path_merge"

    @staticmethod
    def get_config():
        config = {}
        if os.path.exists(FileManager.CONFIG_FILE):
            with open(FileManager.CONFIG_FILE, "r") as file:
                for line in file:
                    key, value = line.strip().split("=")
                    config[key] = value
        return config

    @staticmethod
    def set_config(key, value):
        config = FileManager.get_config()
        config[key] = value
        with open(FileManager.CONFIG_FILE, "w") as file:
            for k, v in config.items():
                file.write(f"{k}={v}\n")

    @staticmethod
    def set_base_path_kw():
        path = filedialog.askdirectory()
        if path:
            FileManager.set_config(FileManager.BASE_PATH_KW, path)
            Messagebox.show_info("Ścieżka zapisana", "Sukces")

    @staticmethod
    def get_base_path_kw():
        return FileManager.get_config().get(FileManager.BASE_PATH_KW)

    @staticmethod
    def set_base_path_merge():
        path = filedialog.askdirectory()
        if path:
            FileManager.set_config(FileManager.BASE_PATH_MERGE, path)
            Messagebox.show_info("Ścieżka zapisana", "Sukces")

    @staticmethod
    def get_base_path_merge():
        return FileManager.get_config().get(FileManager.BASE_PATH_MERGE)

    @staticmethod
    def get_files_list(base_path):
        if not base_path or not os.path.exists(base_path):
            Messagebox.show_error( "Nie ustawiono poprawnej ściezki bazowej", "Błąd",)
            return []

        files = glob.glob(os.path.join(base_path, "*"))
        files.sort(key=os.path.getmtime, reverse=True)
        return files

    @staticmethod
    def get_email_receiver():
        emails = FileManager.get_config().get(FileManager.EMAIL_RECEIVERS)
        return emails

    @staticmethod
    def set_email_receiver():
        emails = simpledialog.askstring("Odbiorcy e-mail",
                                        "Podaj adresy e-mail oddzielone średnikiem:",
                                        initialvalue=FileManager.get_email_receiver() or "")
        if emails:
            FileManager.set_config(FileManager.EMAIL_RECEIVERS, emails)
            Messagebox.show_info("Adresy zapisane", "Sukces")

    @staticmethod
    def load_excel(file_path):
        try:
            df = pd.read_excel(file_path, engine="openpyxl")
            return df
        except Exception as e:
            Messagebox.show_error(f"Błąd podczas wczytywania pliku: {e}", "Błąd")
            return None
