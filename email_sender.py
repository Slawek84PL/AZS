import win32com.client
from tkinter import messagebox

from premailer import transform


class EmailSender:

    @staticmethod
    def send_email(html_body, excel_file, suma, emailTo):
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")

            mail = outlook.CreateItem(0)
            mail.Display()
            default_signature = mail.HTMLBody if mail.HTMLBody else ""

            mail.To = emailTo if emailTo else ""
            mail.Subject = f"{excel_file} GYAL. Suma ładunków {suma}"
            mail.HTMLBody = html_body + default_signature

            mail.Display()
        except Exception as e:
            messagebox.showerror("Błąd", f"Nie udało się wysłać maila : {str(e)}")

    @staticmethod
    def build_html_body(pivot_table_countries, pivot_table_cities):
        html_pivot_countries = pivot_table_countries.to_html(
            classes="excel") if pivot_table_countries is not None else ""
        html_pivot_cities = pivot_table_cities.to_html(classes="excel") if pivot_table_cities is not None else ""

        style = """
            <style type="text/css">
                table.excel {
                    border-collapse: collapse;
                    font-family: Arial, sans-serif;
                    font-size: 12px;
                }
                table.excel th, table.excel td {
                    border: 1px solid #000;
                    padding-top: 1px;
                    padding-bottom: 1px;
                    padding-left: 4px;
                    padding-right: 4px;
                    text-align: center;
                }
                table.excel th {
                    background-color: MediumSeaGreen;
                }
            </style>
            """

        html_body = f"""
            <html>
                <head>
                    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
                    {style}
                </head>
                <body>
                    <p>Poniżej zestawienie ilości ładunków z podziałem na kraje i logistykę</p>
                    {html_pivot_countries}
                    <br>
                    <p>Poniżej z podziałem na magazyny i logistykę</p>
                    {html_pivot_cities}
                </body>
            </html>
            """
        return transform(html_body)
