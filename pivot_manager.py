import pandas as pd


class PivotManager:

    @staticmethod
    def create_pivot_table(df, pivot_index, pivot_columns, pivot_values):
        try:
            for col in [pivot_values]:
                if col not in df.columns:
                    print(f"Błąd, prawdopodobnie wczytano niepoprawny plik")
                    return None

            pivot_table = pd.pivot_table(
                df,
                index=pivot_index,
                columns=pivot_columns,
                values=pivot_values,
                aggfunc='count',
                fill_value=' ',
                margins=True,
                margins_name='Suma'
            )

            return pivot_table
        except Exception as e:
            print(f"Błąd przy tworzeniu tabeli przestawnej: {e}")

    @staticmethod
    def fix_column_name(df, pivot_index_countries, pivot_index_cities, pivot_columns):
        new_columns = list(df.columns)
        new_columns[0] = pivot_index_countries
        new_columns[1] = pivot_index_cities
        new_columns[3] = pivot_columns
        df.columns = new_columns
        return df

    @staticmethod
    def get_suma(pivot_table_cities):
        try:
            return pivot_table_cities.loc["Suma", "Suma"]
        except KeyError:
            return 0
