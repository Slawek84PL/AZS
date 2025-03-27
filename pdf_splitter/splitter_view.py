import os
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

        self.style.configure("Treeview", font=("Arial", 15), rowheight=50)
        self.style.configure("Treeview.Heading", font=("Arial", 14, "bold"))

        self.create_config_section()
        self.create_action_buttons()
        self.create_main_layout()
        self.create_status_bar()

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
        preview_frame = Frame(parent)
        preview_frame.pack(side=RIGHT, fill=Y)

        self.preview_canvas = Canvas(preview_frame, width=1200, bg="grey")
        self.preview_canvas.pack(side=LEFT, fill=Y, expand=YES)

        self.scrollbar = Scrollbar(preview_frame, orient=VERTICAL, command=self.preview_canvas.yview)
        self.scrollbar.pack(side=RIGHT, fill=Y)
        self.preview_canvas.configure(yscrollcommand=self.scrollbar.set)

        self.inner_frame = Frame(self.preview_canvas)
        self.inner_frame.bind("<Configure>",
                              lambda e: self.preview_canvas.configure(scrollregion=self.preview_canvas.bbox("all")))
        self.preview_canvas.create_window((0, 0), window=self.inner_frame, anchor="w")
        self.preview_canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _on_mousewheel(self, event):
        self.preview_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def create_group_table(self, parent):
        columns = ["barcode", "title", "pages"]
        self.group_table = Treeview(parent, columns=columns, show=HEADINGS, bootstyle="info", style="Treeview")

        headers = {
            "barcode": "Kod kreskowy",
            "title": "Dostawca",
            "pages": "Strony PDF"
        }

        for col in columns:
            self.group_table.heading(col, text=headers[col])
            self.group_table.column(col, width=200, anchor=W)

        self.group_table.pack(side=LEFT, fill=BOTH, expand=NO)
        self.group_table.bind("<Button-1>", self.on_table_click)

    def create_action_buttons(self):
        button_frame = Frame(self)
        button_frame.pack(pady=10)

        (Button(button_frame, text="Dodaj wiersz", command=self.add_row_popup, bootstyle=(SUCCESS, OUTLINE))
         .pack(side=LEFT,padx=10, ipadx=10,ipady=5))

        Button(button_frame, text="Podziel i zapisz", command=self.split_pdf_action, bootstyle=(SUCCESS, OUTLINE)).pack(
            side=LEFT, padx=10, ipadx=10, ipady=5)

        (Button(button_frame, text="Wczytaj PDF", command=self.load_pdf_dialog, bootstyle=(SUCCESS, OUTLINE))
         .pack(side=LEFT, padx=10,ipadx=10, ipady=5))

        (Button(button_frame, text="Wyczyść", command=self.clear_all, bootstyle=OUTLINE)
         .pack(side=LEFT, padx=10,ipadx=10, ipady=5))

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
        self.preview_images.clear()
        self.selected_pages.clear()
        for widget in self.inner_frame.winfo_children():
            widget.destroy()

        doc = fitz.open(self.pdf_path)
        for i, page in enumerate(doc):
            pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            img_tk = ImageTk.PhotoImage(img)
            self.preview_images.append(img_tk)

            lbl = Label(self.inner_frame, image=img_tk, bg="red", cursor="hand2", highlightbackground="red", highlightthickness=2)
            lbl.image = img_tk
            lbl.pack(pady=5)
            lbl.bind("<Button-1>", lambda e, idx=i: self.toggle_page_selection(idx, e.widget))

        self.update_status()

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

    def on_table_click(self, event):
        row_id = self.group_table.identify_row(event.y)
        if row_id and self.selected_pages:
            current = self.group_table.set(row_id, "pages")
            new_pages = sorted(self.selected_pages)
            new_values = ",".join(str(p + 1) for p in new_pages)
            new_value = current + ("," if current and new_values else "") + new_values
            self.group_table.set(row_id, "pages", new_value)
            self.selected_pages.clear()
            for widget in self.inner_frame.winfo_children():
                widget.config(highlightbackground="red", highlightthickness=2)
            self.update_status()

    def split_pdf_action(self):
        if not self.pdf_path:
            Messagebox.show_error( "Nie załadowano pliku PDF.", "Błąd")
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

        Label(popup, text="Kod kreskowy").grid(row=0, column=0, padx=10, pady=5, sticky=W)
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