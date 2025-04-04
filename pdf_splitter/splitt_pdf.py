import os

import fitz

from file_manager import FileManager


class PDFGenerator:

    @staticmethod
    def generate_splits(original_pdf_path, row_data, output_folder):
        try:
            doc = fitz.open(original_pdf_path)
            files_saved = 1
            pages_saved = len(doc)
            stickers_saved = 0

            for barcode, pages_str in row_data:
                if not barcode or not pages_str:
                    continue

                pages = [int(p.strip()) - 1 for p in pages_str.split(",") if p.strip().isdigit()]
                if not pages:
                    continue

                new_doc = fitz.open()
                for p in pages:
                    if 0 <= p < len(doc):
                        new_doc.insert_pdf(doc, from_page=p, to_page=p)

                out_path = os.path.join(output_folder, f"{barcode}.pdf")
                new_doc.save(out_path)
                stickers_saved += 1

            config = FileManager.get_config()
            prev_files = int(config.get("split_files_count", 0))
            prev_pages = int(config.get("split_pages_count", 0))
            prev_stickers = int(config.get("stickers_saved", 0))

            FileManager.set_config("split_files_count", str(prev_files + files_saved))
            FileManager.set_config("split_pages_count", str(prev_pages + pages_saved))
            FileManager.set_config("saved_stickers", str(prev_stickers + stickers_saved))

            return True, None
        except Exception as e:
            return False, str(e)
