import os

import openpyxl as opx
import pandas as pd

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
            print("No file named " + self.name + " found, creating one...")
            with self.new_pd_excel_writer() as xl_writer:
                df = pd.DataFrame()
                df.to_excel(xl_writer)
            print(self.name + " created.")
        
        self.reload_opx_workbook()
        self.reload_pandas_excel()
    
    def opx_save(self):
        self.opx_workbook.save(self.absolute_path)
    
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
    
    def add_sheet(self, sheet_name):
        if sheet_name not in self.sheet_names:
            with self.new_pd_excel_writer(mode="a", if_sheet_exists="error") as xl_writer:
                df = pd.DataFrame()
                
                df.to_excel(xl_writer, sheet_name=sheet_name)
            self.reload_opx_workbook()
            self.reload_pandas_excel()
        else:
            print(sheet_name + " is already in " + self.name)
        return self.sheets[sheet_name]
    
    def reload_pandas_excel(self):
        self.excel = pd.ExcelFile(self.absolute_path, self.engine)
        self.sheet_names = self.excel.sheet_names
        self.sheets = {}
        for name in self.sheet_names:
            self.sheets[name] = SheetDool(name, self)
    
    def reload_opx_workbook(self):
        self.opx_workbook = opx.load_workbook(self.absolute_path)
    
    def new_pd_excel_writer(self, mode="w", if_sheet_exists=None):
        xl_writer = pd.ExcelWriter(self.absolute_path,
                                   engine=self.engine,
                                   mode=mode,
                                   if_sheet_exists=if_sheet_exists,
                                   )
        return xl_writer

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
    
    def read(self, header=0, index_col=0):
        """Returns a dataframe from a sheet, assuming that the sheet is 
        just a simple table with a header row at the top and the index 
        col on the left"""
        df = pd.read_excel(self.excel,
                           sheet_name=self.name,
                           header=header,
                           index_col=index_col)
        return df
    
    def hide_link_text(self, cell_column, cell_row, replacement="LINK"):
        """If the referenced cell's value is a link, the cell is changed 
        to a hyperlink cell and its value is changed to the 
        replacement text. Nothing is returned. This function does not save 
        the changes, that must be done separately"""
        index_0_row = cell_row + 1
        index_0_column = cell_column + 1
        cell = self.opx_sheet.cell(row=index_0_row, column=index_0_column)
        value = cell.value
        if value is not None and "http" in str(value):
            cell.hyperlink = value
            cell.value = replacement
            cell.style = "Hyperlink"
    
    def hide_colon_reference_links(self, colon_reference, replacement="LINK"):
        """formats all cells that have "http" in them as hyperlinks and 
        changes their values to the replacement text. This function 
        must be used to format links in excel since pandas cannot 
        input links that are larger than 255 characters due to excel's 
        hyperlink function character limit"""
        cells = erm.colon_reference_to_int(colon_reference)
        left_top = cells[0]
        right_bot = cells[1]
        cols = list(range(left_top[0], right_bot[0] + 1))
        rows = list(range(left_top[1], right_bot[1] + 1))
        self.reload_opx_sheet()
        for row in rows:
            for col in cols:
                self.hide_link_text(col, row, replacement)
        self.opx_save()
    
    def _add_dataframe(self, df, start="A1", mode="a", if_sheet_exists=None):
        """Writes a dataframe to the sheet starting from the start cell. 
        Returns the reference range of cells that were written to for 
        any future processing"""
        start_list = erm.reference_to_int(start)
        start_column = start_list[0]
        start_row = start_list[1]
        with self.excel_dool.new_pd_excel_writer(mode=mode,
                                                 if_sheet_exists=if_sheet_exists
                                                ) as xl_writer:
            df.to_excel(xl_writer,
                        sheet_name=self.name,
                        startrow=start_row,
                        startcol=start_column,
                        )
        bottom_right = [start_column + df.shape[1], start_row + df.shape[0]]
        return start + ":" + erm.int_to_reference(bottom_right)
    
    def write_dataframe(self, df, start="A1"):
        return self._add_dataframe(df, start=start, mode="w")
    
    def append_dataframe(self, df, start="A1", if_sheet_exists="replace"):
        return self._add_dataframe(df, start=start, mode="a",
                            if_sheet_exists=if_sheet_exists)
    
    def auto_fit_columns(self):
        """Iterates through all columns and sizes them to fit their 
        contents. Doesn't return anything"""
        self.reload_opx_sheet()
        # Iterate over all columns and adjust their widths
        for column in self.opx_sheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 4)
            self.opx_sheet.column_dimensions[column_letter].width = adjusted_width
        self.opx_save()
    
    def reload_opx_sheet(self):
        self.excel_dool.reload_opx_workbook()
        self.opx_sheet = self.excel_dool.opx_workbook[self.name]
    
    def opx_save(self):
        self.excel_dool.opx_save()
    
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

