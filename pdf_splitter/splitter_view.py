import os
from tkinter.ttk import Checkbutton

import fitz
from tkinter import filedialog, StringVar, Label
from ttkbootstrap import Frame, Button, Labelframe, Treeview, Scrollbar, Canvas
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs import Messagebox
import ttkbootstrap as ttk
from PIL import Image, ImageTk

from file_manager import FileManager
from pdf_splitter.splitt_pdf import PDFGenerator


class PDFSplitterView(ttk.Toplevel):
    def __init__(self):
        super().__init__()
        self.title("Dzielenie plików PDF")
        self.geometry("1600x1200")

        self.pdf_path = None
        self.selected_pages = set()
        self.preview_images = []
        self.status_var = StringVar()
        self.checkbox_vars = {}

        self.style.configure("Treeview", font=("Arial", 12), rowheight=30)
        self.style.configure("Treeview.Heading", background="#4a7abc", foreground="white", font=("Arial", 14, "bold"))

        self.create_config_section()
        self.create_action_buttons()
        self.create_main_layout()
        self.create_status_bar()

        self.bind_all("<Control-v>", self.on_paste_clipboard)

    def create_config_section(self):
        config_frame = Labelframe(self, text="Konfiguracja", padding=10)
        config_frame.pack(fill=X, padx=10, pady=10)

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

    def change_input_path(self):
        path = filedialog.askdirectory()
        if path:
            FileManager.set_config("pdf_input_path", path)
            self.input_path_lbl.config(text=path)

    def change_output_path(self):
        path = filedialog.askdirectory()
        if path:
            FileManager.set_config("pdf_output_path", path)
            self.output_path_lbl.config(text=path)

    def create_main_layout(self):
        layout_frame = Frame(self)
        layout_frame.pack(fill=BOTH, expand=YES, padx=10, pady=10)

        self.create_group_table(layout_frame)
        self.create_preview_section(layout_frame)

    def create_preview_section(self, parent):
        # 🔲 Główna ramka podglądu
        preview_frame = Frame(parent)
        preview_frame.pack(fill=BOTH, expand=YES)

        # 🔲 Lewa kolumna – miniatury PDF (w grid)
        self.thumbnail_frame = Frame(preview_frame)
        self.thumbnail_frame.pack(side=LEFT, fill=Y, padx=10)

        # 🔲 Prawa kolumna – duży podgląd strony
        right_preview_frame = Frame(preview_frame)
        right_preview_frame.pack(side=LEFT, fill=BOTH, expand=YES)

        self.preview_canvas = Canvas(right_preview_frame, bg="white")
        self.scrollbar = Scrollbar(right_preview_frame, orient="vertical", command=self.preview_canvas.yview)
        self.preview_canvas.configure(yscrollcommand=self.scrollbar.set)

        self.preview_canvas.pack(side=LEFT, fill=BOTH, expand=YES)
        self.scrollbar.pack(side=RIGHT, fill=Y)

        self.right_image_id = None  # do aktualizacji obrazu

    def _on_mousewheel(self, event):
        self.preview_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

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

    def create_action_buttons(self):
        button_frame = Frame(self)
        button_frame.pack(pady=10)

        (Button(button_frame, text="Dodaj wiersz", command=self.add_row_popup, bootstyle=(SUCCESS, OUTLINE))
         .pack(side=LEFT, padx=10, ipadx=10, ipady=5))

        (Button(button_frame, text="Podziel i zapisz", command=self.split_pdf_action, bootstyle=SUCCESS).pack(
            side=LEFT, padx=10, ipadx=10, ipady=5))

        (Button(button_frame, text="Wczytaj PDF", command=self.load_pdf_dialog, bootstyle=(SUCCESS))
         .pack(side=LEFT, padx=10, ipadx=10, ipady=5))

        (Button(button_frame, text="Wyczyść", command=self.clear_all, bootstyle=OUTLINE)
         .pack(side=LEFT, padx=10, ipadx=10, ipady=5))

    def create_status_bar(self):
        Label(self, textvariable=self.status_var, anchor=W).pack(fill=X, side=BOTTOM, ipady=2)
        self.update_status()

    def update_status(self):
        if not self.selected_pages:
            self.status_var.set("Brak zaznaczonych stron")
        else:
            pages = ", ".join(str(p + 1) for p in sorted(self.selected_pages))
            self.status_var.set(f"Zaznaczone strony: {pages}")

    def load_pdf_dialog(self):
        initial_path = FileManager.get_config().get("pdf_input_path", "")
        path = filedialog.askopenfilename(
            filetypes=[("Pliki PDF", "*.pdf")],
            initialdir=initial_path
        )
        if path:
            self.pdf_path = path
            self.load_pdf_preview()

    def load_pdf_preview(self):
        self.selected_pages.clear()
        self.checkbox_vars = {}
        for widget in self.thumbnail_frame.winfo_children():
            widget.destroy()

        self.pdf_doc = fitz.open(self.pdf_path)
        cols = 2
        row, col = 0, 0

        for i, page in enumerate(self.pdf_doc):
            pix = page.get_pixmap(matrix=fitz.Matrix(0.5, 0.5))
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            img.thumbnail((180, 250))
            img_tk = ImageTk.PhotoImage(img)

            var = ttk.BooleanVar()
            self.checkbox_vars[i] = var

            frame = Frame(self.thumbnail_frame, padding=5)
            frame.grid(row=row, column=col, padx=5, pady=5)

            lbl = Label(frame, image=img_tk, bg="white", cursor="hand2")
            lbl.image = img_tk
            lbl.pack()
            lbl.bind("<Button-1>", lambda e, idx=i: self.show_page_preview(idx))

            chk = Checkbutton(frame, variable=var, text=f"{i + 1}")
            chk.pack()

            col += 1
            if col >= cols:
                col = 0
                row += 1

        self.update_status()

    def show_page_preview(self, page_index):
        if not self.pdf_doc:
            return

        page = self.pdf_doc.load_page(page_index)
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))  # Wysoka jakość
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        img_tk = ImageTk.PhotoImage(img)

        # 🔄 Wyczyść stary obraz
        self.preview_canvas.delete("all")

        # 🖼️ Wstaw nowy obraz
        self.preview_canvas.image = img_tk  # ważne!
        self.right_image_id = self.preview_canvas.create_image(0, 0, anchor=NW, image=img_tk)

        # 📏 Ustaw scroll i rozmiar obszaru
        self.preview_canvas.config(scrollregion=(0, 0, pix.width, pix.height))

    def toggle_page_selection(self, page_index, widget):
        if page_index in self.selected_pages:
            self.selected_pages.remove(page_index)
            widget.config(highlightbackground="red", highlightthickness=2)
        else:
            self.selected_pages.add(page_index)
            widget.config(highlightbackground="green", highlightthickness=2)
        self.update_status()

    def clear_all(self):
        self.selected_pages.clear()
        self.preview_images.clear()
        for widget in self.inner_frame.winfo_children():
            widget.destroy()
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
            Messagebox.show_error("Błąd wklejania", str(e))

    def on_table_click(self, event):
        row_id = self.group_table.identify_row(event.y)
        if row_id:
            # ⬇️ Pobierz listę stron zaznaczonych checkboxami
            selected = [i for i, var in self.checkbox_vars.items() if var.get()]

            if selected:
                current = self.group_table.set(row_id, "pages")
                new_values = ",".join(str(p + 1) for p in sorted(selected))  # +1 dla ludzkiej numeracji
                new_value = current + ("," if current and new_values else "") + new_values

                # 📝 Ustaw nowe wartości
                self.group_table.set(row_id, "pages", new_value)

                # ♻️ Wyczyść zaznaczenia checkboxów
                for i in selected:
                    self.checkbox_vars[i].set(False)

                self.update_status()

    def split_pdf_action(self):
        if not self.pdf_path:
            Messagebox.show_error("Nie załadowano pliku PDF.", "Błąd")
            return

        output_dir = FileManager.get_config().get("pdf_output_path") or os.getcwd()

        row_data = [
            (self.group_table.set(row_id, "barcode"), self.group_table.set(row_id, "pages"))
            for row_id in self.group_table.get_children()
        ]

        success, error = PDFGenerator.generate_splits(self.pdf_path, row_data, output_dir)
        if success:
            Messagebox.show_info("Pliki PDF zostały zapisane.", "Sukces")
        else:
            Messagebox.show_error(error, "Błąd")

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
