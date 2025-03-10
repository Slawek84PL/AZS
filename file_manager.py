import glob
import os.path
from tkinter import filedialog, messagebox
import pandas as pd


class FileManager:
    CONFIG_FILE = "config.txt"

    @staticmethod
    def get_base_path():
        if os.path.exists(FileManager.CONFIG_FILE):
            with open(FileManager.CONFIG_FILE, "r") as file:
                return file.read().strip()
        return None

    @staticmethod
    def set_base_path():
        path = filedialog.askdirectory()
        if path:
            with open(FileManager.CONFIG_FILE, "w") as file:
                file.write(path)
            messagebox.showinfo("Sukces", "Ścieżka zapisana")

    @staticmethod
    def get_files():
        base_path = FileManager.get_base_path()
        if not base_path or not os.path.exists(base_path):
            messagebox.showerror("Błąd", "Nie ustawiono poprawnej ściezki bazowej")
            return []

        files = glob.glob(os.path.join(base_path, "*"))
        files.sort(key=os.path.getmtime, reverse=True)
        return files

    @staticmethod
    def load_excel(file_path):
        try:
            df = pd.read_excel(file_path, engine="openpyxl")
            return df
        except Exception as e:
            print(f"Błąd podczas wczytywania pliku: {e}")
            return None
