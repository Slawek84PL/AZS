import os
import tkinter as tk

import ttkbootstrap as ttk
from more_itertools import side_effect
from ttkbootstrap import OUTLINE
from ttkbootstrap.dialogs import Messagebox
from ttkbootstrap.constants import *
from ttkbootstrap import utility

from email_sender import EmailSender
from file_manager import FileManager
from pivot_manager import PivotManager


class SendKw(ttk.Frame):

    def __init__(self, master):
        super().__init__(master, padding=15)
        self.pack(fill=BOTH, expand=YES)

        self.buttons_lf = ttk.Labelframe(self, text="Konfiguracja", padding=15)
        self.buttons_lf.pack(fill=X, expand=YES, anchor=N)

        self.create_base_path()
        self.create_email_receiver()
        self.create_results_view()

    def create_base_path(self):
        path_row = ttk.Frame(self.buttons_lf)
        path_row.pack(fill=X, expand=YES, pady=5)
        path_lbl = ttk.Label(path_row, text="Folder KW", width=15)
        path_lbl.pack(side=LEFT, padx=(15, 0))
        base_path = FileManager.get_base_path()
        print(base_path)
        path_ent = ttk.Label(path_row, text=base_path)
        path_ent.pack(side=LEFT, fill=X, expand=YES, padx=1)
        search_btn = ttk.Button(
            master=path_row,
            text="Zmień",
            command=FileManager.set_base_path,
            bootstyle=OUTLINE,
            width=8
        )
        search_btn.pack(side=LEFT, padx=5)

    def create_email_receiver(self):
        path_row = ttk.Frame(self.buttons_lf)
        path_row.pack(fill=X, expand=YES, pady=5)
        path_lbl = ttk.Label(path_row, text="Odbiorcy", width=15)
        path_lbl.pack(side=LEFT, padx=(15, 0))
        base_path = FileManager.get_email_receiver()
        print(base_path)
        path_ent = ttk.Label(path_row, text=base_path)
        path_ent.pack(side=LEFT, fill=X, expand=YES, padx=1)
        search_btn = ttk.Button(
            master=path_row,
            text="Zmień",
            command=FileManager.set_email_receiver,
            bootstyle=OUTLINE,
            width=8
        )
        search_btn.pack(side=LEFT, padx=5)

    def create_results_view(self):
        self.resultview = ttk.Treeview(
            master=self,
            bootstyle=INFO,
            columns=[0],
            show=HEADINGS
        )
        self.resultview.pack(fill=BOTH, expand=YES, pady=10)

        self.resultview.heading(0,text="Nazwa", anchor=W)

        self.resultview.column(
            column=0,
            anchor=W,
            width=utility.scale_size(self, 550),
            stretch=False
        )

        file_list = FileManager.get_files_list()
        for file in file_list:
            self.resultview.insert("", index=END, values=(os.path.basename(file),))

class FileEmailApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Raporty")
        self.root.geometry('400x300')

        style = ttk.Style("minty")

        ttk.Label(root, text="Wybierz rozwiązanie", font=("Arial", 12, "bold")).pack(pady=10)
        ttk.Button(root, text="Wyślij KW Gyal", command=self.show_file_list).pack(pady=10)
        # ttk.Button(root, text="Podziel Analizy", command=self.show_file_list).pack(pady=10)

    def show_file_list(self):
        file_window = ttk.Toplevel(self.root)
        file_window.title("pliki")
        file_window.geometry("400x600")

        ttk.Label(file_window, text=FileManager.get_base_path()).pack(pady=10)
        ttk.Button(file_window, text="Ustaw ścieżkę do plików", command=FileManager.set_base_path).pack(pady=10)

        ttk.Label(file_window, text=FileManager.get_email_receiver()).pack(pady=10)
        ttk.Button(file_window, text="Ustaw odbiorców adresów", command=FileManager.set_email_receiver).pack(pady=10)

        file_list = FileManager.get_files_list()

        listbox = tk.Listbox(file_window, width=80, height=20)
        listbox.pack(pady=10)

        for file in file_list:
            listbox.insert(tk.END, os.path.basename(file))

        if file_list:
            ttk.Button(file_window, text="Wyślij zestawienie transportów",
                       command=lambda: self.process_selected_file(listbox, file_list)).pack(
                pady=10)

    def process_selected_file(self, listbox, file_list):
        selected_index = listbox.curselection()
        if not selected_index:
            Messagebox.show_error("Błąd", "Nie wybrano pliku")
            return

        file_path = file_list[selected_index[0]]
        print(file_path)
        df = FileManager.load_excel(file_path)

        if df is None or df.empty:
            Messagebox.show_error("Błąd", "Plik nie zawiera danych!")
            return

        pivot_index_countries = ['Kraj']
        pivot_index_cities = ['Magazyn']
        pivot_columns = ['Logistyka']
        pivot_values = 'Order no'

        df = PivotManager.fix_column_name(df, pivot_index_countries[0], pivot_index_cities[0], pivot_columns[0])

        pivot_table_countries = PivotManager.create_pivot_table(df, pivot_index_countries, pivot_columns, pivot_values)
        if pivot_table_countries is None:
            Messagebox.show_error("Błąd", "Prawdopodobnie próbujesz wysłać raport z niepoprawnego pliku.")
            return

        pivot_table_cities = PivotManager.create_pivot_table(df, pivot_index_cities, pivot_columns, pivot_values)
        print(pivot_table_countries)
        print(pivot_table_cities)

        html_body = EmailSender.build_html_body(pivot_table_countries, pivot_table_cities)

        file_name = os.path.basename(file_path)[:5]
        suma = PivotManager.get_suma(pivot_table_cities)

        if suma == 0:
            Messagebox.show_error("Błąd", "Suma transportów jest niepoprawna. Nie wysłano maila")
            return

        emailTo = FileManager.get_email_receiver()
        print(emailTo)

        EmailSender.send_email(html_body, file_name, suma, emailTo)


if __name__ == '__main__':
    app = ttk.Window("Automatyzacje", "journal")
    # app = FileEmailApp(root)
    SendKw(app)
    app.mainloop()
