import os.path
import pandas as pd
import win32com
from openpyxl import load_workbook

from file_manager import FileManager
from ttkbootstrap.dialogs import Messagebox


class FileMerger():

    @staticmethod
    def read_files(first_file_name, second_file_name):
        base_path = FileManager.get_base_path_merge()
        first_file_path = os.path.join(base_path, first_file_name)
        second_file_path = os.path.join(base_path, second_file_name)

        first_df = FileManager.load_excel(first_file_path)
        second_df = FileManager.load_excel(second_file_path)

        if first_df is None or second_df is None:
            Messagebox.show_error("Jeden z wybranych plików nie posiada zawartości")
            return

        FileMerger.proceed_df(first_df, second_df, second_file_path)

    @staticmethod
    def proceed_df(first_df, second_df, second_file_path):
        first_df = FileMerger.remove_header(first_df)

        merged_df = pd.concat([second_df, first_df], ignore_index=True)

        merged_df = FileMerger.replace_data(merged_df)

        merged_df = FileMerger.remove_header(merged_df)

        FileMerger.update_file(merged_df, second_file_path)

    @staticmethod
    def update_file(merged_df, second_file_path):
        wb = load_workbook(second_file_path)
        ws = wb["ZAM"]

        ws.delete_rows(2, ws.max_row - 1)

        for r_idx, row in enumerate(merged_df.values, start=2):
            for c_idx, value in enumerate(row, start=1):
                ws.cell(row=r_idx, column=c_idx, value=value)

        FileMerger.run_vba_macro(second_file_path)

    @staticmethod
    def run_vba_macro(second_file_path):
        macro_name = "ListaZaladunkowa"

        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False

        workbook = excel.Workbooks.Open(second_file_path)

        try:
            workbook.Application.Run(f"'{workbook.Name}'!{macro_name}")
            FileMerger.copy_sheet_to_new_workbook(workbook, "LoadingList")

        except Exception as e:
            Messagebox.show_error(f"Błąd podczas wykonywania makra: {e}", "Błąd")

        workbook.Close(SaveChanges=True)
        excel.Quit()

    @staticmethod
    def copy_sheet_to_new_workbook(workbook, sheet_name):
        output_file = os.path.join(FileManager.get_base_path_merge(), "ll.xlsx")

        try:
            sheet = workbook.Sheets(sheet_name)
            new_workbook = workbook.Application.Workbooks.Add()
            sheet.Copy(Before=new_workbook.Sheets(1))

            new_workbook.SaveAs(output_file)
            print(f"Arkusz {sheet_name} skopiowany do {output_file}")

            new_workbook.Close(SaveChanges=True)
            return output_file

        except Exception as e:
            Messagebox.show_error(f"Błąd podczas kopiowania arkusza: {e}", "Błąd")
            return None

    @staticmethod
    def remove_header(df):
        df = df.iloc[0:]
        return df

    @staticmethod
    def replace_data(merged_df):
        column_name = merged_df.columns[9]
        merged_df[column_name] = merged_df[column_name].str.replace("KOD", "KON")
        merged_df[column_name] = merged_df[column_name].str.replace("GYD", "GYA")

        return merged_df
