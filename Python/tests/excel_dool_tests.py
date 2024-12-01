import sys
import os

import pandas as pd
sys.path.append("../src")

from excel_dool import ExcelDool

#from website_table import WebsiteTable
#from bs4 import BeautifulSoup

"""
# Testing Reading Excel Files
test_file = "test_xlsx.xlsx"
xl = ExcelDool(test_file)
sheet = xl.sheets["Sheet1"]
"""

# Testing Writing Excel Files
test_write_file = "test_out_xlsx.xlsx"

if os.path.exists(test_write_file):
    os.remove(test_write_file)

xl_out = ExcelDool(test_write_file)



data = [dict(name='Google', url='http://www.google.com'),
        dict(name='Stackoverflow', url='http://stackoverflow.com')]
df = pd.DataFrame(data)

out_sheet = xl_out.sheets["Sheet1"]
colon_ref = out_sheet.write_dataframe(df, "B2")
out_sheet.hide_colon_reference_links(colon_ref)


"""
with ExcelDool(test_write_file, mode="w") as xl_dool:
    data = [dict(name='Google', url='http://www.google.com'),
            dict(name='Stackoverflow', url='http://stackoverflow.com')]
    df = pd.DataFrame(data)
    df.to_excel(xl_dool.writer)
"""


"""
with ExcelDool(test_write_file, mode="w") as xl_dool:
    xl_dool.sheets["Sheet1"].format_link_text(1,1)
"""