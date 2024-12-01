import sys
import os

sys.path.append("../src")

from job_parser import JobParser
from excel_dool import ExcelDool

ignore_file_name = "C:/Users/George/Documents/trunk/References/Job Search/_Searching/ignored_jobs.json"
jobs_file_name = "C:/Users/George/Documents/trunk/References/Job Search/_Searching/all_jobs.json"
companies = "C:/Users/George/Documents/trunk/References/Job Search/_Searching/company_jobs"
#jp = JobParser(jobs_file_name, companies, ignore_file_name)
jp = JobParser(jobs_file_name, companies)
out_file = "C:/Users/George/Documents/trunk/References/Job Search/_Searching/unignored_jobs.json"
jp.write_to_json(out_file)
df = jp.create_excel_dataframe()

excel_output = "C:/Users/George/Documents/trunk/References/Job Search/_Searching/job_status.xlsx"

#if os.path.exists(excel_output):
#    os.remove(excel_output)

xl = ExcelDool(excel_output)

sheet_name = "Sheet1"
xl.add_sheet(sheet_name)
xl.sheets[sheet_name].append_dataframe(df)
xl.sheets[sheet_name].auto_fit_columns()

sheet_name = "Sheet2"
xl.add_sheet(sheet_name)
output_range = xl.sheets[sheet_name].append_dataframe(df)
xl.sheets[sheet_name].hide_colon_reference_links(output_range)
xl.sheets[sheet_name].auto_fit_columns()

xl.reload_pandas_excel()

print(xl.sheets["Sheet1"].read())