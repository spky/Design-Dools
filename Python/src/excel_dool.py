import os
import re
import math

import openpyxl as opx
import pandas as pd
import numpy as np

class ExcelDool:
    """
    
    Instance Variables:
    engine -- see keyword arguments
    absolute_path -- the absolute file path of the given file
    name -- The name and extension of the given file
    folder -- the path of the folder containing the given file
    mode -- read, write, or append mode for pandas
    writer -- ExcelWriter pandas object for the excel file
    excel -- a pandas ExcelFile object used to parse sheets into dataframes
    opx_workbook --  openpyxl workbook for the excel file
    sheet_names -- list of the names of the sheets in the excel file
    sheets = dictionary with sheet names + SheetDool objects value pairs
    
    Keyword arguments:
    mode -- chooses whether the excel is just to be read, written to or 
    appended to
    engine -- choice of engine to read an excel-like file (default openpyxl. 
    Options: xlrd, odf, pyxlsb, calamine)
    """
    
    def __init__(self, file_path, mode="r", engine="openpyxl"):
        self.engine = engine
        self.absolute_path = os.path.abspath(file_path)
        self.name = os.path.basename(file_path)
        self.folder = os.path.dirname(self.absolute_path)
        self.mode = mode
        self.writer = None
        
        # if file does not exist, create a blank excel file
        if not os.path.exists(self.absolute_path):
            print("No file named " + self.name + " found, creating one")
            with pd.ExcelWriter(self.absolute_path,
                                engine=self.engine,
                                mode="w") as new_xl_writer:
                df = pd.DataFrame()
                df.to_excel(new_xl_writer)
        
        self.excel = pd.ExcelFile(self.absolute_path, self.engine)
        self.opx_workbook = opx.load_workbook(self.absolute_path)
        self.sheet_names = self.excel.sheet_names
        self.sheets = {}
        for name in self.sheet_names:
            self.sheets[name] = SheetDool(name, self)
    
    def __enter__(self):
        match self.mode:
            case "r":
                self.writer = None
            case "w":
                self.writer = pd.ExcelWriter(self.absolute_path,
                                             engine=self.engine,
                                             mode=self.mode)
            case "a":
                self.writer = pd.ExcelWriter(self.absolute_path,
                                             engine=self.engine,
                                             mode=self.mode)
            case _:
                self.writer = None
        return self
    
    def __exit__(self, file_path, engine="openpyxl", mode="r"):
        if self.writer is not None:
            self.writer.close()

class SheetDool:
    
    def __init__(self, sheet_name, excel_dool):
        self.name = sheet_name
        self.excel_dool = excel_dool
        self.excel = self.excel_dool.excel
        self.opx_sheet = self.excel_dool.opx_workbook[self.name]
        self.table_names = self.opx_sheet.tables.keys()
    
    def read_colon_reference(self, colon_reference):
        references = colon_reference_to_int(colon_reference)
        
        columns = list(range(references[0][0], references[1][0]+1))
        skip_rows = references[0][1]
        number_of_rows = references[1][1] - references[0][1]
        return pd.read_excel(self.excel, 
                             sheet_name=self.name, 
                             skiprows=skip_rows, 
                             engine=self.engine, 
                             usecols=columns, 
                             nrows=number_of_rows)
    
    def read_table(self, table_name):
        if table_name in self.table_names:
            table = self.opx_sheet.tables[table_name]
            return self.read_colon_reference(table.ref)
        else:
            raise ValueError(self.name 
                            + " does not have a table named " 
                            + table_name)
    
    def format_link_text(self, cell_column, cell_row):
        
        cell = self.opx_sheet.cell(row=cell_row, column=cell_column)
        print(cell.value)
    
    @property
    def engine(self):
        return self.excel.engine
    
    @property
    def absolute_path(self):
        return self.excel.absolute_path
    
    @property
    def file_name(self):
        return self.excel.name
    
    @property
    def folder(self):
        return self.excel.folder

def colon_reference_to_int(colon_reference):
    """Splits a string of format [A-Z][1-9]:[A-Z][1-9], converts them to 
    integer lists, and returns a 2 element list of those lists"""
    references = colon_reference.split(":")
    if len(references) == 2:
        left_top = reference_to_int(references[0])
        right_bottom = reference_to_int(references[1])
    else:
        raise ValueError("Reference not in a format of (ref):(ref)")
    return [left_top, right_bottom]

def reference_to_int(reference):
    """Returns a list of [column, row], indexed at 0"""
    parsed = parse_reference(reference)
    if parsed is not None:
        column = column_to_int(parsed[0])
        row = int(parsed[1]) - 1
        return [column, row]
    else:
        return None

def parse_reference(reference):
    """Checks if the string matches a regex for an excel reference. 
    Returns a list consisting of the column and row references if it does 
    match. If it doesn't match it just returns None"""
    regex = re.compile("^(?P<column>[a-zA-Z]+)(?P<row>[0-9]+)$")
    match = re.match(regex, reference)
    if match is not None:
        return [match["column"], match["row"]]
    else:
        return None

def int_to_reference(int_reference_list):
    """Returns the equivalent reference to integer defined, indexed-0 column and row
    return int_to_column(column) + str(row)"""
    if int_reference_list[1] < 1:
        raise ValueError("Row cannot be less than 1")
    return int_to_column(int_reference_list[0]) + str(int_reference_list[1])

def column_to_int(column_reference):
    """Returns column number corresponding to letter reference, ex: a or A = 0, 
    AA = 26. Letter case type is ignored.
    """
    base = 0
    integer = -1
    for letter in reversed(list(column_reference)):
        lower_letter = letter.lower()
        integer += (ord(lower_letter) - 96)*(26**base)
        base += 1
    return integer

def int_to_column(integer):
    """Returns the column reference letter, ex: 0 = A, 26 == AA. Always returns 
    capital letters"""
    # The alphabetical column system is a base-26 number system
    BASE = 26
    ASCII_OFFSET = 96
    column = ""
    # Initialize int_quotient and modulus so the loop will run at least once
    int_quotient = -1
    modulus = 0
    while integer:
        int_quotient = math.floor(integer/BASE)
        modulus = integer % BASE
        if int_quotient:
            column += chr(int_quotient + ASCII_OFFSET).upper()
            integer = modulus
        else:
            integer = 0
    column += chr(modulus + ASCII_OFFSET + 1).upper()
    return column