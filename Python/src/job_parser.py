import os
import json
import pandas as pd

class JobParser:
    
    def __init__(self, jobs_file, company_directory, ignore_file=None):
        
        self.jobs_file = jobs_file
        self.ignore_file = ignore_file
        self.jobs = {}
        self.load_in_jobs(company_directory)
        
        if ignore_file is not None:
            self.load_in_ignored(ignore_file)
    
    def load_in_ignored(self, ignore_file):
        with open(ignore_file, "r") as file:
            self.ignored = json.load(file)
            self.remove_ignored_jobs()
    
    def remove_ignored_jobs(self, ignore=None):
        if ignore is None:
            ignore = self.ignored
        
        for company in ignore:
            for job in ignore[company]:
                if job in self.jobs[company]:
                    self.jobs[company].pop(job)
    
    def load_in_jobs(self, company_directory):
        files = os.listdir(company_directory)
        
        for file in files:
            if file.endswith(".json"):
                company_name = file.removesuffix(".json")
                company_name = company_name.replace("_", " ")
                
                filepath = company_directory + "/" + file
                with open(filepath, "r") as company_file:
                    self.jobs[company_name] = json.load(company_file)
    
    def create_excel_dataframe(self):
        data = []
        for company in self.jobs:
            for job in self.jobs[company]:
                job_dict = dict(company_name=company,
                                name=job,
                                url=self.jobs[company][job]
                                )
                data.append(job_dict)
        return pd.DataFrame(data)
    
    def write_ignore_file(self, file_name, indent=4):
        with open(file_name, "w") as file:
            json.dump(self.ignored, file, indent=indent)
    
    def write_to_json(self, file_name, indent=4):
        with open(file_name, "w") as file:
            json.dump(self.jobs, file, indent=indent)
        print(file_name + " written")