import pandas as pd
import logging
from abc import ABC, abstractmethod
import re

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class DataLoader(ABC):
    def __init__(self, file_path: str):
        self.file_path = file_path

    @abstractmethod
    def load_data(self):
        pass

class ExcelDataLoader(DataLoader):
    def __init__(self, file_path: str, possible_sheets: list):
        super().__init__(file_path)
        self.possible_sheets = possible_sheets

    def find_valid_sheet_name(self):
        """ Check for the existence of possible sheet names and return the first valid one. """
        xl = pd.ExcelFile(self.file_path)
        for sheet_name in self.possible_sheets:
            if sheet_name in xl.sheet_names:
                logging.info(f"Found valid sheet name: {sheet_name}")
                return sheet_name
        raise ValueError("No valid sheet name found in the list of possible sheets.")

    def find_header_row(self, sheet_name) -> int:
        df = pd.read_excel(self.file_path, sheet_name, header=None)
        columns_of_interest = ["Reference", "Referencja", "Description", "Ordered in Std Pack", "Unit", "Loop Size"]
        regex_pattern = '|'.join([re.escape(col) for col in columns_of_interest])

        for i in range(len(df)):
            if df.iloc[i].astype(str).str.contains(regex_pattern, regex=True, na=False).any():
                return i
        raise ValueError("Header row not found")

    def load_data(self):
        sheet_name = self.find_valid_sheet_name()
        header_row = self.find_header_row(sheet_name)
        df = pd.read_excel(self.file_path, sheet_name, header=[header_row, header_row+1])
        logging.info(f"Data loaded successfully from sheet '{sheet_name}' with header row at: {header_row}")
        return df

class DataFrameProcessor:
    def __init__(self, dataframe: pd.DataFrame):
        self.df = dataframe

    def normalize_headers(self):
        """ Normalize headers by flattening tuples and cleaning them up. """
        self.df.columns = [' '.join(col).replace('\n', ' ').replace('  ', ' ').strip() for col in self.df.columns]
        logging.info(f"Normalized headers: {self.df.columns.tolist()}")

    def rename_and_select_columns(self, columns_mapping: dict):
        """ Map and filter DataFrame based on required columns only, handling multiple potential names per column. """
        self.df = self.df.copy()
        actual_columns = {col: col for col in self.df.columns}
        selected_columns = {}
        for col_patterns, new_name in columns_mapping.items():
            matched_col = next((actual_col for pattern in col_patterns for actual_col in actual_columns if pattern in actual_col), None)
            if matched_col:
                selected_columns[matched_col] = new_name
            else:
                logging.error(f"No matching columns found for patterns: {col_patterns}")
                raise KeyError(f"No matching columns found for patterns: {col_patterns}")
        self.df = self.df[list(selected_columns.keys())]
        self.df.rename(columns=selected_columns, inplace=True)

    def add_additional_columns(self):
        """ Add additional columns and compute 'Ilość' based on 'Ordered in Std Pack' and 'Loop Size'. """
        self.df['Ordered in Std Pack'] = self.df['Ordered in Std Pack'].fillna(0)
        self.df['Loop Size'] = self.df['Loop Size'].fillna(0)
        self.df['Sztuk'] = (self.df['Ordered in Std Pack'].astype(float) * self.df['Loop Size'].astype(float)).astype(int)
        additional_info = ["Klient", "Oczekiwany termin realizacji", "Termin potwierdzony",
                           "Uwagi dla wszystkich", "Uwagi niewidoczne dla produkcji", 
                           "Atrybut 1 (opcjonalnie)", "Atrybut 2 (opcjonalnie)", "Atrybut 3 (opcjonalnie)"]
        for col in additional_info:
            self.df[col] = ""

    def filter_data(self):
        """ Filter out rows where 'Ordered in Std Pack' is zero. """
        self.df = self.df[self.df['Ordered in Std Pack'].astype(float) != 0]

    def remove_empty_rows(self):
        """ Remove rows where all values are empty. """
        self.df.replace("", float("NaN"), inplace=True)
        self.df.dropna(how='all', inplace=True)

    def remove_weird_rows(self):
        """ Remove rows with specific weird content. """
        weird_content = ["Transportation mode", "Supplier Contact signature", "for the promise", "____________________"]
        self.df = self.df[~self.df.apply(lambda row: any(weird in str(row) for weird in weird_content), axis=1)]

    def finalize_columns_order(self):
        """ Finalize the order of columns to match the desired output CSV format. """
        final_columns_order = ["Klient", "Oczekiwany termin realizacji", "Termin potwierdzony",
                               "Zewn. nr zamówienia", "Produkt", "Sztuk",
                               "Uwagi dla wszystkich", "Uwagi niewidoczne dla produkcji",
                               "Atrybut 1 (opcjonalnie)", "Atrybut 2 (opcjonalnie)", "Atrybut 3 (opcjonalnie)"]
        self.df = self.df[final_columns_order]