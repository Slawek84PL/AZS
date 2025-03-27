import os
from tkinter import filedialog, Label as TkLabel, StringVar

import ttkbootstrap as  ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
from tkinter.constants import *
from ttkbootstrap import Frame, Button, Labelframe, Treeview, OUTLINE, SUCCESS, utility, Canvas, Style, Scrollbar, Y, \
    RIGHT, LEFT, BOTH, YES, W, CENTER, Toplevel, HEADINGS, INFO, X, BOTTOM
from ttkbootstrap.dialogs import Messagebox
from PIL import Image, ImageTk
import fitz

from file_manager import FileManager
from pdf_splitter.splitt_pdf import PDFGenerator


class PDFSplitterView(TkinterDnD.Tk):

    def __init__(self):
        super().__init__()
        self.title("Dzielenie plików PDF")
        self.style = Style("journal")
        self.geometry("1200x900")

        self.pdf_path = None
        self.groups = {}
        self.selected_pages = set()
        self.preview_images = []
        self.status_var = StringVar()

        self.create_config_section()
        self.create_main_layout()
        self.create_action_buttons()
        self.create_status_bar()

        self.bind_all("<Control-v>", self.on_paste_clipboard)

    def create_config_section(self):
        config_frame = Labelframe(self, text="Konfiguracja", padding=10)
        config_frame.pack(fill=X, padx=10, pady=10)

        output_row = Frame(config_frame)
        output_row.pack(fill=X, pady=5)
        TkLabel(output_row, text="Folder PDF (wyjściowy):", width=25).pack(side=LEFT, padx=5)
        self.output_path_lbl = TkLabel(output_row, text=FileManager.get_config().get("pdf_output_path", ""))
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

        self.preview_canvas = Canvas(preview_frame, width=400, bg="white")
        self.preview_canvas.pack(side=LEFT, fill=Y, expand=YES)
        self.preview_canvas.drop_target_register(DND_FILES)
        self.preview_canvas.dnd_bind('<<Drop>>', self.on_pdf_drop)

        self.scrollbar = Scrollbar(preview_frame, orient="vertical", command=self.preview_canvas.yview)
        self.scrollbar.pack(side=RIGHT, fill=Y)
        self.preview_canvas.configure(yscrollcommand=self.scrollbar.set)

        self.inner_frame = Frame(self.preview_canvas)
        self.inner_frame.bind("<Configure>",
                              lambda e: self.preview_canvas.configure(scrollregion=self.preview_canvas.bbox("all")))
        self.preview_canvas.create_window((0, 0), window=self.inner_frame, anchor="nw")

        # Scroll mouse wheel
        self.preview_canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _on_mousewheel(self, event):
        self.preview_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def create_group_table(self, parent):
        style = ttk.Style()
        style.configure("Custom.Treeview.Heading",
                        font=("Segoe UI", 14, "bold"),
                        background="#f0f0f0",
                        foreground="#333333",
                        padding=10)

        style.configure("Custom.Treeview",
                        font=("Segoe UI", 13),
                        rowheight=30,  # Zwiększona wysokość
                        background="white",  # Białe tło
                        foreground="black",  # Czarny tekst
                        borderwidth=5,  # Cienka ramka
                        relief=FLAT)

        columns = ["barcode", "title", "pages"]
        self.group_table = Treeview(parent, columns=columns, show=HEADINGS, bootstyle=INFO, style="Custom.Treeview")

        headers = {
            "barcode": "Kod kreskowy",
            "title": "Dostawca",
            "pages": "Strony PDF"
        }

        for col in columns:
            self.group_table.heading(col, text=headers[col])
            self.group_table.column(col, width=200 if col != "pages" else 250, anchor=W)

        self.group_table.pack(side=LEFT, fill=BOTH, expand=YES)
        self.group_table.bind("<Button-1>", self.on_table_click)

    def create_action_buttons(self):
        button_frame = Frame(self)
        button_frame.pack(pady=10)

        Button(button_frame, text="Podziel PDF i zapisz", bootstyle=(SUCCESS, OUTLINE),
               command=self.split_pdf_action).pack(side=LEFT, padx=10, ipadx=10, ipady=5)
        Button(button_frame, text="Wyczyść zaznaczenia", bootstyle=OUTLINE, command=self.clear_all).pack(side=LEFT,
                                                                                                         padx=10,
                                                                                                         ipadx=10,
                                                                                                         ipady=5)

    def create_status_bar(self):
        self.status_label = TkLabel(self, textvariable=self.status_var, anchor=W)
        self.status_label.pack(fill=X, side=BOTTOM, ipady=2)
        self.update_status()

    def update_status(self):
        if not self.selected_pages:
            self.status_var.set("Brak zaznaczonych stron")
        else:
            pages = ", ".join(str(p + 1) for p in sorted(self.selected_pages))
            self.status_var.set(f"Zaznaczone strony: {pages}")

    def on_pdf_drop(self, event):
        path = event.data.strip('{}')
        if path.lower().endswith(".pdf") and os.path.isfile(path):
            self.pdf_path = path
            self.load_pdf_preview()
        else:
            Messagebox.show_error("Przeciągnięty plik nie jest PDF-em.", "Błąd", )

    def load_pdf_preview(self):
        self.preview_images.clear()
        self.selected_pages.clear()
        for widget in self.inner_frame.winfo_children():
            widget.destroy()

        doc = fitz.open(self.pdf_path)
        for i, page in enumerate(doc):
            pix = page.get_pixmap(matrix=fitz.Matrix(0.5, 0.5))
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            img_tk = ImageTk.PhotoImage(img)
            self.preview_images.append(img_tk)

            lbl = TkLabel(self.inner_frame, image=img_tk, bg="white", cursor="hand2")
            lbl.pack(pady=5)
            lbl.bind("<Button-1>", lambda e, idx=i: self.toggle_page_selection(idx, e.widget))

        self.update_status()

    def toggle_page_selection(self, page_index, widget):
        if page_index in self.selected_pages:
            self.selected_pages.remove(page_index)
            widget.config(highlightthickness=0)
        else:
            self.selected_pages.add(page_index)
            widget.config(highlightbackground="blue", highlightthickness=2)
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
                widget.config(highlightthickness=0)
            self.update_status()

    def split_pdf_action(self):
        if not self.pdf_path:
            Messagebox.show_error("Nie załadowano pliku PDF.", "Błąd", )
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


if __name__ == '__main__':
    app = PDFSplitterView()
    app.mainloop()
