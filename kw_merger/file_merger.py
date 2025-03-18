import os.path

import pandas as pd
from ttkbootstrap.dialogs import Messagebox

from file_manager import FileManager


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

        FileMerger.proceed_df(first_df, second_df, first_file_name)

    @staticmethod
    def proceed_df(first_df, second_df, first_file_name):

        merged_df = pd.concat([second_df, first_df], ignore_index=True)

        merged_df = FileMerger.replace_data(merged_df)

        report_df = FileMerger.transform_and_group_orders(merged_df)

        FileManager.save_excel_file(report_df, first_file_name[:5] + " LL")



    @staticmethod
    def replace_data(merged_df):
        column_name = merged_df.columns[9]
        merged_df[column_name] = merged_df[column_name].str.replace("KOD", "KON")
        merged_df[column_name] = merged_df[column_name].str.replace("GYD", "GYA")

        return merged_df

    @staticmethod
    def transform_and_group_orders(df):

        country_map = {
            "SK": 1,
            "CZ": 2,
            "PL": 3,
            "HU": 4,
            "RS": 5,
            "BG": 6,
            # Dodasz resztę kodów
        }

        result_rows = []

        for order, group in df.groupby("Rendelési szám"):
            main_location = group[group["Infó"] == "HAVI"]
            main_quantity = main_location["Raklapok \nszáma"].sum() if not main_location.empty else 0

            additional_locations = group[group["Infó"] != "HAVI"]

            if not additional_locations.empty:
                grouped_additional = (
                    additional_locations.groupby("Infó", as_index=False)["Raklapok \nszáma"]
                    .sum()
                )

                additional_text = " ".join(
                    f"{row['Infó']}({row['Raklapok \nszáma']}pal)"
                    for _, row in grouped_additional.iterrows()
                )
            else:
                additional_text = ""

            # Transformacje kolumn
            group["B"] = group["A címzett \nraktára"]
            group["A"] = group["A fogadó \nországa"].map(country_map).astype(str) + " " + group["A fogadó \nországa"]

            result_rows.append([
                group["A"].iloc[0],
                group["B"].iloc[0],
                "",
                "",
                main_quantity,
                additional_text,
                # additional_locations,
                order,
            ])

        result_df = pd.DataFrame(result_rows, columns=[
            "A fogadó ", "A címzett","dod1", "dod2", "HAVI", "Comment", "Order no",
        ])

        return result_df
