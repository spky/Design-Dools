import sys
import os
import math

import pandas as pd
sys.path.append("../src")

from excel_dool import (ExcelDool, 
                        column_to_int,
                        int_to_column,
                        int_to_reference,
                        parse_reference,
                        reference_to_int)

from website_table import WebsiteTable
from bs4 import BeautifulSoup

"""
# Testing Reading Excel Files
test_file = "test_xlsx.xlsx"
xl = ExcelDool(test_file)
"""

# Testing Writing Excel Files
test_write_file = "test_out_xlsx.xlsx"

if os.path.exists(test_write_file):
    os.remove(test_write_file)

def make_clickable(val):
    # target _blank to open new window
    return '<a target="_blank" href="{}">{}</a>'.format(val, val)


with ExcelDool(test_write_file, mode="w") as xl_dool:
    data = [dict(name='Google', url='http://www.google.com'),
            dict(name='Stackoverflow', url='http://stackoverflow.com')]
    df = pd.DataFrame(data)
    df.to_excel(xl_dool.writer)

with ExcelDool(test_write_file, mode="w") as xl_dool:
    xl_dool.sheets["Sheet1"].format_link_text(1,1)