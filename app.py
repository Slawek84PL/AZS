import os

import ttkbootstrap as ttk
from ttkbootstrap import utility
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs import Messagebox

from email_sender import EmailSender
from file_manager import FileManager
from kw_merger.view import ViewMerger
from pivot_manager import PivotManager


class SendKw(ttk.Toplevel):

    def __init__(self, master):
        super().__init__(master, resizable=(False, False))

        self.title("Wyślij raport KW")
        self.buttons_lf = ttk.Labelframe(self, text="Konfiguracja", padding=15)
        self.buttons_lf.pack(fill=X, expand=YES, anchor=N, padx=10)

        self.create_base_path()
        self.create_email_receiver()
        self.create_results_view()
        self.create_send_btn()

    def create_base_path(self):
        path_row = ttk.Frame(self.buttons_lf)
        path_row.pack(fill=X, expand=YES, pady=5)
        path_lbl = ttk.Label(path_row, text="Folder KW", width=15)
        path_lbl.pack(side=LEFT, padx=(15, 0))
        base_path = FileManager.get_base_path_kw()
        path_ent = ttk.Label(path_row, text=base_path)
        path_ent.pack(side=LEFT, fill=X, expand=YES, padx=1)
        search_btn = ttk.Button(
            master=path_row,
            text="Zmień",
            command=FileManager.set_base_path_kw,
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
        self.resultview.pack(fill=BOTH, expand=YES, pady=10, padx=10)

        self.resultview.heading(0, text="Nazwa", anchor=W)

        self.resultview.column(
            column=0,
            anchor=W,
            width=utility.scale_size(self, 550),
            stretch=False
        )

        file_list = FileManager.get_files_list(FileManager.get_base_path_kw())
        for file in file_list:
            self.resultview.insert("", index=END, values=(os.path.basename(file),))

    def create_send_btn(self):
        self.send_btn = ttk.Button(
            master=self,
            text="Wyślij raport",
            command=self.process_selected_file,
            bootstyle=(SUCCESS, OUTLINE)
        )
        self.send_btn.pack(pady=20, ipadx=10, ipady=5)

    def process_selected_file(self):
        selected_index = self.resultview.selection()
        print(selected_index)
        if not selected_index:
            Messagebox.show_error("Nie wybrano pliku", "Błąd", )
            return

        file_name = self.resultview.item(selected_index[0], "values")[0]
        file_path = os.path.join(FileManager.get_base_path_kw(), file_name)
        # print(file_path)
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
        print(suma)
        print(file_name)
        print(emailTo)

        EmailSender.send_email(html_body, file_name, suma, emailTo)


class MainApp(ttk.Window):
    def __init__(self):
        super().__init__("Automatyzacje", "journal")
        self.geometry("400x400")

        ttk.Button(
            self,
            text="Wyślij raport KW Gyal",
            command=self.open_send_kw,
            bootstyle=SUCCESS
        ).pack(pady=50, ipadx=10, ipady=10)

        ttk.Button(
            self,
            text="Scal pliki KW Gyal",
            command=self.open_merge_kw,
            bootstyle=SUCCESS
        ).pack(pady=50, ipadx=10, ipady=10)

    def open_send_kw(self):
        SendKw(self)

    def open_merge_kw(self):
        ViewMerger(self)


if __name__ == '__main__':
    app = MainApp()
    app.mainloop()
