# Spreadsheet Column Unique Value Extractor
# Made by Patrick Dugan, inspired by a tedious spreadsheet task.
import argparse 
import pandas as pd
import pathlib
import sys 
from itertools import chain

class scuve():
    def __init__(self):
        BANNER = """
/======================================\\
||  ____   ____  _   ___     _______  ||
|| / ___| / ___|| | | \ \   / / ____| ||
|| \___ \| |    | | | |\ \ / /|  _|   ||
||  ___) | |___ | |_| | \ V / | |___  ||
|| |____(_)____(_)___(_) \_(_)|_____| ||
\======================================/
Spreadsheet.Column.Unique.Val.Extractor
"""
        print(BANNER)
        self.args = scuve.parse_args()


    @staticmethod
    def parse_args():
        description  = """
Extracts unique values from the specified columns. 
Supports extracting values from cells with string formatted lists.
"""
        parser = argparse.ArgumentParser(description=description)
        parser.add_argument('-f','--file_path', help="File path of spreadsheet. Supports CSV, TSV & XLSX", type=str, required=True)
        parser.add_argument('-c','--column_names', help='String or CSV string of column names', type=str)
        args = parser.parse_args()
        return args
    
    def validate_and_load(self):
        file_type = pathlib.Path(self.args.file_path).suffix
        match file_type:
            case '.csv':
                data = pd.read_csv(self.args.file_path)
            case '.tsv':
                data =pd.read_table(self.args.file_path)
            case '.xlsx':
                data = pd.read_excel(self.args.file_path)
            case _ :
                sys.exit(f"File Type {file_type} not supported.")
        if self.args.column_names is None:
            print(f"Extracting column names from: '{self.args.file_path}'")
            for index, column in enumerate(data.columns):
                print(f"{index}: {column}")
            columns = input("\nEnter the column number(s) to extract from (comma separated).\n").split(",")
            columns = [column.strip() for column in columns]
            for column in columns:
                try:
                    data.columns[int(column)]
                except ValueError as error:
                    print(f"Invalid column number: {column}")
                except IndexError as error:
                    print(f"Column number {column} does not exist.") 
            self.args.column_names = [data.columns[int(column)] for column in columns]
        else:
            self.args.column_names = self.args.column_names.split(",")
            self.args.column_names = [name.strip() for name in self.args.column_names]
            for name in self.args.column_names:
                try:
                    data.loc[:,name]
                except KeyError as error:
                    print(f"{name} not in column headers.")
        return data       

    def extract(self):
        data = self.validate_and_load()
        extracted_vals = list(chain.from_iterable([list(data[name].unique()) for name in self.args.column_names]))
        extracted_vals = [str(val).strip('"').strip("'").strip() for val in extracted_vals]
        tmp_list = []
        for val in extracted_vals:
            if val.startswith('['):
                extracted_vals.remove(val)
                for extracted_val in val.split(","):
                    tmp_list.append(extracted_val.strip('[').strip(']'))            
        for val in tmp_list:
            extracted_vals.append(val)
        extracted_vals = list(set([val.strip("'").strip('"').strip() for val in extracted_vals]))           
        print(f"Unique Values Extracted:\n{', '.join(extracted_vals)}")


if __name__ == '__main__':
    scuve_ = scuve()
    scuve_.extract()