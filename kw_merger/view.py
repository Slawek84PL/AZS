import os
from tkinter.constants import *

import ttkbootstrap as ttk
from ttkbootstrap import OUTLINE, SUCCESS, INFO, HEADINGS, utility

from file_manager import FileManager
from kw_merger.file_merger import FileMerger


class ViewMerger(ttk.Toplevel):

    def __init__(self, master):
        super().__init__(master, resizable=(False, False))

        self.title("Scalanie plików KW")
        self.buttons_lf = ttk.Labelframe(self, text="Konfiguracja", padding=15)
        self.buttons_lf.pack(fill=X, expand=YES, anchor=N, padx=10, pady=10)

        self.create_base_path()
        self.create_results_view()
        self.create_merge_btn()

    def create_base_path(self):
        path_row = ttk.Frame(self.buttons_lf)
        path_row.pack(fill=X, expand=YES, pady=5)
        path_lbl = ttk.Label(path_row, text="Folder KW", width=15)
        path_lbl.pack(side=LEFT, padx=(15, 0))
        base_path = FileManager.get_base_path_merge()
        path_ent = ttk.Label(path_row, text=base_path)
        path_ent.pack(side=LEFT, fill=X, expand=YES, padx=1)
        search_btn = ttk.Button(
            master=path_row,
            text="Zmień",
            command=FileManager.set_base_path_merge,
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

        file_list = FileManager.get_files_list(FileManager.get_base_path_merge())
        for file in file_list:
            self.resultview.insert("", index=END, values=(os.path.basename(file),))

    def create_merge_btn(self):
        self.send_btn = ttk.Button(
            master=self,
            text="Scal pliki",
            command=self.proceed_selected_files,
            bootstyle=(SUCCESS, OUTLINE)
        )
        self.send_btn.pack(pady=20, ipadx=10, ipady=5)

    def proceed_selected_files(self):
        files = self.resultview.selection()
        if len(files) == 2:
            first_file_name = self.resultview.item(files[0], "values")[0]
            second_file_name = self.resultview.item(files[1], "values")[0]
            FileMerger.read_files(first_file_name, second_file_name)
