import os
from tkinter import filedialog, StringVar, Label
import pdf_splitter.splitter_helper as helper

import fitz
import ttkbootstrap as ttk
from ttkbootstrap import Frame, Button, Labelframe, Treeview, Scrollbar, Canvas
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs import Messagebox

from file_manager import FileManager
from pdf_splitter.splitt_pdf import PDFGenerator


class PDFSplitterView(ttk.Toplevel):
    def __init__(self):
        super().__init__()
        # self.iconbitmap(os.path.join(os.path.dirname(__file__), "jas_pdf.ico"))
        self.title("Dzielenie plików PDF")
        self.state("zoomed")

        self.pdf_path = None
        self.selected_pages = set()
        self.status_bar = StringVar()

        self.style.configure("Treeview", font=("Arial", 12), rowheight=30)
        self.style.configure("Treeview.Heading", background="#4a7abc", foreground="white", font=("Arial", 14, "bold"))

        self.create_config_section()
        self.create_action_buttons()
        self.create_main_layout()
        self.create_status_bar()

        self.bind_all("<Control-v>", self.on_paste_clipboard)

    def create_config_section(self):
        config_frame = Labelframe(self, text="Konfiguracja", padding=5)
        config_frame.pack(fill=X, padx=5, pady=5)

        input_row = Frame(config_frame)
        input_row.pack(fill=X, pady=5)
        Label(input_row, text="Folder z dokumentami:", width=25).pack(side=LEFT, padx=5)
        self.input_path_lbl = Label(input_row, text=FileManager.get_config().get("pdf_input_path", ""))
        self.input_path_lbl.pack(side=LEFT, fill=X, expand=YES)
        Button(input_row, text="Zmień", command=self.change_input_path, bootstyle=OUTLINE).pack(side=LEFT, padx=5)

        output_row = Frame(config_frame)
        output_row.pack(fill=X, pady=5)
        Label(output_row, text="Folder send:", width=25).pack(side=LEFT, padx=5)
        self.output_path_lbl = Label(output_row, text=FileManager.get_config().get("pdf_output_path", ""))
        self.output_path_lbl.pack(side=LEFT, fill=X, expand=YES)
        Button(output_row, text="Zmień", command=self.change_output_path, bootstyle=OUTLINE).pack(side=LEFT, padx=5)

    def create_action_buttons(self):
        button_frame = Frame(self)
        button_frame.pack(pady=10)

        (Button(button_frame, text="Edytuj strony", command=self.edit_pages, bootstyle=INFO)
         .pack(side=LEFT, padx=10, ipadx=10, ipady=5))
        (Button(button_frame, text="Wklej dane ze schowka", command=self.on_paste_clipboard, bootstyle=SECONDARY)
         .pack(side=LEFT, padx=10, ipadx=10, ipady=5))
        (Button(button_frame, text=f"Dodaj wiersz", command=self.add_row_popup,
                bootstyle=(SUCCESS, OUTLINE))
         .pack(side=LEFT, padx=10, ipadx=10, ipady=5))
        (Button(button_frame, text="Podziel i zapisz", command=self.split_pdf_action, bootstyle=SUCCESS)
         .pack(side=LEFT, padx=10, ipadx=10, ipady=5))
        (Button(button_frame, text="Wczytaj PDF", command=self.load_pdf_dialog, bootstyle=SUCCESS)
         .pack(side=LEFT, padx=10, ipadx=10, ipady=5))
        (Button(button_frame, text="Wyczyść", command=self.clear_all, bootstyle=DANGER)
         .pack(side=RIGHT, padx=10, ipadx=10, ipady=5))

    def create_main_layout(self):
        layout_frame = Frame(self)
        layout_frame.pack(fill=BOTH, expand=YES, padx=10, pady=10)

        self.create_group_table(layout_frame)
        self.create_preview_section(layout_frame)

    def create_group_table(self, parent):
        columns = ["barcode", "title", "pages"]
        self.group_table = Treeview(parent, columns=columns, show=HEADINGS, bootstyle="info", style="Treeview")

        self.group_table.heading("barcode", text="Nazwa pliku")
        self.group_table.heading("title", text="Dostawca")
        self.group_table.heading("pages", text="Strony PDF")

        self.group_table.column("barcode", width=175, anchor=W)
        self.group_table.column("title", width=275, anchor=W)
        self.group_table.column("pages", width=150, anchor=CENTER)

        self.group_table.pack(side=LEFT, fill=BOTH, expand=NO)
        self.group_table.bind("<Button-1>", self.on_table_click)

    def edit_pages(self):
        from tkinter import Toplevel, Entry
        preview = Toplevel(self)
        preview.state("zoomed")
        preview.title("Podgląd i edycja")

        canvas = Canvas(preview, bg="white")
        scrollbar = Scrollbar(preview, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=LEFT, fill=BOTH, expand=YES)
        scrollbar.pack(side=RIGHT, fill=Y)

        scrollable_frame = Frame(canvas)
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        scrollable_frame.bind("<Enter>", lambda e: self.bind_mousewheel(scrollable_frame))
        scrollable_frame.bind("<Leave>", lambda e: self.unbind_mousewheel(scrollable_frame))

        self.active_row = self.group_table.selection()
        self.page_numbers = self.group_table.item(self.active_row, "values")[2].split(",")

        cols, x, y = helper.get_resolution(len(self.page_numbers))
        row, col = 0, 0

        self.pdf_doc = fitz.open(self.pdf_path)
        for page_index in self.page_numbers:
            page = self.pdf_doc.load_page(int(page_index) - 1)
            img_tk = helper.build_image_from_page(page, x, y, 1)

            frame = Frame(scrollable_frame, padding=2)
            frame.grid(row=row, column=col, padx=2, pady=2)

            lbl = Label(frame, image=img_tk, bg="white", cursor="hand2")
            lbl.image = img_tk
            lbl.pack()

            Button(frame, text="Usuń", command=lambda f=frame, p=page_index: self.delete_page(f, p),
                   bootstyle=DANGER).pack(pady=5)

            col += 1
            if col >= cols:
                col = 0
                row += 1

    def delete_page(self, frame, page_index):
        frame.destroy()
        if page_index in self.page_numbers:
            self.page_numbers.remove(page_index)
        self._update_pages()

    def _update_pages(self):
        if self.active_row:
            row_id = self.active_row[0]
            new_value = ",".join(self.page_numbers)
            self.group_table.set(row_id, "pages", new_value)

    def create_preview_section(self, parent):
        preview_frame = Frame(parent)
        preview_frame.pack(fill=BOTH, expand=YES)

        left_frame = Frame(preview_frame)
        left_frame.pack(side=LEFT, fill=Y)

        self.thumb_canvas = Canvas(left_frame, bg="white", width=400)
        self.thumb_scrollbar = Scrollbar(left_frame, orient="vertical", command=self.thumb_canvas.yview)
        self.thumb_canvas.configure(yscrollcommand=self.thumb_scrollbar.set)

        self.thumb_canvas.pack(side=LEFT, fill=Y, expand=YES)
        self.thumb_scrollbar.pack(side=RIGHT, fill=Y)

        self.thumbnail_frame = Frame(self.thumb_canvas)
        self.thumb_canvas_window = self.thumb_canvas.create_window((0, 0), window=self.thumbnail_frame, anchor="nw")

        self.thumbnail_frame.bind("<Configure>",
                                  lambda e: self.thumb_canvas.configure(scrollregion=self.thumb_canvas.bbox("all")))
        self.thumb_canvas.bind("<Enter>", lambda e: self.bind_mousewheel(self.thumb_canvas))
        self.thumb_canvas.bind("<Leave>", lambda e: self.unbind_mousewheel(self.thumb_canvas))

        right_preview_frame = Frame(preview_frame)
        right_preview_frame.pack(side=LEFT, fill=BOTH, expand=YES)

        self.preview_canvas = Canvas(right_preview_frame, bg="white")
        self.scrollbar = Scrollbar(right_preview_frame, orient="vertical", command=self.preview_canvas.yview)
        self.preview_canvas.configure(yscrollcommand=self.scrollbar.set)

        self.preview_canvas.pack(side=LEFT, fill=BOTH, expand=YES)
        self.scrollbar.pack(side=RIGHT, fill=Y)

        self.preview_canvas.bind("<Enter>", lambda e: self.bind_mousewheel(self.preview_canvas))
        self.preview_canvas.bind("<Leave>", lambda e: self.unbind_mousewheel(self.preview_canvas))

    def create_status_bar(self):
        Label(self, textvariable=self.status_bar, anchor=W).pack(fill=X, side=BOTTOM, ipady=2)
        self.update_status()

    def load_pdf_dialog(self):
        initial_path = FileManager.get_config().get("pdf_input_path", "")
        path = filedialog.askopenfilename(filetypes=[("Pliki PDF", "*.pdf")], initialdir=initial_path, parent=self)
        if path:
            self.pdf_path = path
            self.load_pdf_preview()

    def load_pdf_preview(self):
        self.clear_thumbnails()
        self.clear_preview()

        self.pdf_doc = fitz.open(self.pdf_path)
        cols = 2
        row, col = 0, 0

        for i, page in enumerate(self.pdf_doc):
            img_tk = helper.build_image_from_page(page, 200, 250, 0.5)

            frame = Frame(self.thumbnail_frame, padding=2)
            frame.grid(row=row, column=col, padx=2, pady=2)

            lbl = Label(frame, image=img_tk, bg="white", cursor="hand2")
            lbl.image = img_tk
            lbl.pack()
            lbl.bind("<Button-1>", lambda idx=i, index=i, widget=lbl: self.toggle_page_selection(index, widget=widget))

            col += 1
            if col >= cols:
                col = 0
                row += 1

        self.thumb_canvas.yview_moveto(0)
        self.update_status()

    def clear_thumbnails(self):
        for widget in self.thumbnail_frame.winfo_children():
            widget.destroy()

    def clear_preview(self):
        self.preview_canvas.delete("all")
        self.preview_canvas.config(scrollregion=(0, 0, 0, 0))

    def show_page_preview(self, page_index):
        if not self.pdf_doc:
            return
        page = self.pdf_doc.load_page(page_index)
        img_tk = helper.build_image_from_page(page, 900, 1050, 1.5)

        self.preview_canvas.delete("all")
        self.preview_canvas.image = img_tk
        self.preview_canvas.create_image(0, 0, anchor=NW, image=img_tk)
        self.preview_canvas.config(scrollregion=(0, 0, 1000, 1050))

    def toggle_page_selection(self, page_index, widget):
        if page_index in self.selected_pages:
            self.selected_pages.remove(page_index)
            self.clear_preview()
            widget.config(highlightbackground="red", highlightthickness=2)
        else:
            self.selected_pages.add(page_index)
            widget.config(highlightbackground="#90EE90", highlightthickness=3)
            self.show_page_preview(page_index)
        self.update_status()

    def clear_all(self):
        self.selected_pages.clear()
        self.clear_thumbnails()
        self.clear_preview()
        for row in self.group_table.get_children():
            self.group_table.delete(row)
        self.pdf_path = None
        self.update_status()

    def on_paste_clipboard(self, event=None):
        try:
            raw = self.clipboard_get()
            rows = raw.strip().split("\n")
            for row in rows:
                parts = row.strip().split("\t")
                if len(parts) >= 4:
                    title = parts[1]
                    barcode = parts[3]
                    self.group_table.insert("", END, values=(barcode, title, ""))
        except Exception as e:
            Messagebox.show_error("Błąd wklejania", str(e), parent=self)

    def on_table_click(self, event):
        row_id = self.group_table.identify_row(event.y)
        if row_id:
            selected = [i for i in self.selected_pages]
            if selected:
                current = self.group_table.set(row_id, "pages")
                new_values = ",".join(str(p + 1) for p in sorted(selected))
                new_value = current + ("," if current and new_values else "") + new_values
                self.group_table.set(row_id, "pages", new_value)

        for i, frame in enumerate(self.thumbnail_frame.winfo_children()):
            label = frame.winfo_children()[0]  # to jest Label z miniaturą
            if i in self.selected_pages:
                label.config(highlightbackground="orange", highlightthickness=2)

        self.clear_preview()
        self.selected_pages.clear()
        self.update_status()

    def split_pdf_action(self):
        if not self.pdf_path:
            Messagebox.show_error("Nie załadowano pliku PDF.", "Błąd", parent=self)
            return

        output_dir = FileManager.get_config().get("pdf_output_path") or os.getcwd()
        row_data = [
            (self.group_table.set(row_id, "barcode"), self.group_table.set(row_id, "pages"))
            for row_id in self.group_table.get_children()
        ]

        success, error = PDFGenerator.generate_splits(self.pdf_path, row_data, output_dir)
        if success:
            Messagebox.show_info("Pliki PDF zostały zapisane.", "Sukces", parent=self)
        else:
            Messagebox.show_error(error, "Błąd", parent=self)

    def add_row_popup(self):
        from tkinter import Toplevel, Entry

        popup = Toplevel(self)
        popup.title("Nowy wiersz")

        Label(popup, text="Nazwa pliku").grid(row=0, column=0, padx=10, pady=5, sticky=W)
        barcode_entry = Entry(popup, width=40)
        barcode_entry.grid(row=0, column=1, padx=10, pady=5)

        Label(popup, text="Dostawca").grid(row=1, column=0, padx=10, pady=5, sticky=W)
        title_entry = Entry(popup, width=40)
        title_entry.grid(row=1, column=1, padx=10, pady=5)

        Label(popup, text="Strony PDF").grid(row=2, column=0, padx=10, pady=5, sticky=W)
        pages_entry = Entry(popup, width=40)
        pages_entry.grid(row=2, column=1, padx=10, pady=5)

        def add_row():
            self.group_table.insert("", "end", values=(barcode_entry.get(), title_entry.get(), pages_entry.get()))
            popup.destroy()

        Button(popup, text="Dodaj", command=add_row, bootstyle=SUCCESS).grid(row=3, column=1, pady=10)

    def change_input_path(self):
        path = filedialog.askdirectory(parent=self)
        if path:
            FileManager.set_config("pdf_input_path", path)
            self.input_path_lbl.config(text=path)

    def change_output_path(self):
        path = filedialog.askdirectory(parent=self)
        if path:
            FileManager.set_config("pdf_output_path", path)
            self.output_path_lbl.config(text=path)

    def update_status(self):
        selected = [i + 1 for i in self.selected_pages]
        if not selected:
            self.status_bar.set("Brak zaznaczonych stron")
        else:
            self.status_bar.set(f"Zaznaczone strony: {', '.join(map(str, selected))}")

    def bind_mousewheel(self, widget):
        self.unbind_mousewheel()
        self.bind_all("<MouseWheel>", lambda e: widget.yview_scroll(int(-1 * (e.delta / 120)), "units"))

    def unbind_mousewheel(self, *_):
        self.unbind_all("<MouseWheel>")
