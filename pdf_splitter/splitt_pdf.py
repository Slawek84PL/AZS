import os

import fitz


class PDFGenerator:

    @staticmethod
    def generate_splits(original_pdf_path, row_data, output_folder):
        try:
            doc = fitz.open(original_pdf_path)
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

            return True, None
        except Exception as e:
            return False, str(e)
