import os
import re
import math

import openpyxl as opx
import pandas as pd
import numpy as np

import excel_reference_managers as erm

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
        references = erm.colon_reference_to_int(colon_reference)
        
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

