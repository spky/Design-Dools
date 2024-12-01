import sys
import time
import re
import json

# Requests documentation: 
# https://requests.readthedocs.io/en/latest/
import requests

# Beautiful Soup documentation: 
# https://beautiful-soup-4.readthedocs.io/en/latest/#
from bs4 import BeautifulSoup

import feedparser

# Selenium documentation:
# https://selenium-python.readthedocs.io/
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.firefox.options import Options as firefox_options

class Webber:
    
    HTML_PARSER = "html.parser"
    
    def __init__(self, url, implementation="selenium", 
                 wait_time=None, browser="firefox", headless=False, 
                 request_data=None):
        self.output_dict = {}
        self.implementation = implementation
        if self.implementation == "selenium":
            self.initialize_driver(browser, headless, wait_time)
            self.update_selenium_page(url)
        elif self.implementation == "requests":
            self.initialize_request(url, request_data)
        elif self.implementation == "post":
            self.url = url
            self.request_data = request_data
            self.update_post()
        elif self.implementation == "rss":
            self.rss = feedparser.parse(url)
            
    def __del__(self):
        if self.implementation == "selenium":
            self.shut_down()
    
    def initialize_driver(self, browser, headless, wait_time):
        if browser == "chrome":
            self.driver = webdriver.Chrome()
        elif browser == "firefox":
            options = firefox_options()
            if headless:
                options.add_argument("--headless")
            self.driver = webdriver.Firefox(options=options)
        
        if wait_time is not None:
            self.driver.implicitly_wait(wait_time)
    
    def initialize_request(self, url, request_data=None):
        self.url = url
        self.request_data = request_data
        self.update_request()
    
    def update_request(self, new_url=None, new_request_data=None):
        if new_url is not None:
            self.url = new_url
        if new_request_data is not None:
            self.request_data = new_request_data
        
        if self.request_data is not None:
            self.request = requests.get(self.url, self.request_data)
        else:
            self.request = requests.get(self.url)
        self.soup = BeautifulSoup(self.request.text, self.HTML_PARSER)
        return self.request
    
    def update_post(self):
        self.request = requests.post(self.url, json=self.request_data)
        self.soup = BeautifulSoup(self.request.text, self.HTML_PARSER)
        return self.request
    
    def update_selenium_page(self, url):
        self.url = url
        self.driver.get(url)
        self.refresh_soup()
    
    def refresh_soup(self):
        self.html = self.driver.page_source
        self.soup = BeautifulSoup(self.html, self.HTML_PARSER)
    
    def search_xpath(self, xpath):
        if self.implementation == "selenium":
            elements = self.driver.find_elements(By.XPATH, xpath)
        else:
            print("requests and bs4 do not support XPATH")
            elements = None
        return elements
    
    def scroll_to_bottom(self, scroll_time=0.5):
        self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(scroll_time)
    
    def click_xpath(self, xpath):
        elements = self.driver.find_elements(By.XPATH, xpath)
        for e in elements:
            self.driver.execute_script("arguments[0].scrollIntoView(true);", e)
            time.sleep(0.5)  # Optional: Wait for the scroll action to complete
            e.click()
        self.html = self.driver.page_source
        self.soup = BeautifulSoup(self.html, self.HTML_PARSER)
    
    def shut_down(self):
        self.driver.quit()
    
    def write_html_to_file(self, file_name):
        with open(file_name, "w", encoding="utf-8") as out_file:
            out_file.write(self.soup.prettify())
        print(file_name + " written")
    
    def output_dict_to_json(self, file_name, indent=4):
        self.write_to_json(file_name, self.output_dict, indent=indent)
        print(file_name + " written")
    
    @staticmethod
    def write_to_json(file_name, content, indent=4):
        with open(file_name, "w") as file:
            json.dump(content, file, indent=indent)
    
    @staticmethod
    def strip_all_white_space(text):
        text = text.strip("\n")
        text = text.strip("\t")
        text = text.strip()
        return text
    
    @staticmethod
    def add_a_to_dict(dictionary, a_element):
        link_name = a_element.string
        link_name = Webber.strip_all_white_space(link_name)
        dictionary[link_name] = a_element.get("href")
    
    @staticmethod
    def write_bs4_file(file_name, element):
        with open(file_name, "w", encoding="utf-8") as out_file:
            out_file.write(element.prettify())
        print(file_name + " written")



# Hard url input

#output_file = "temp_test.html"

def houston_or_webster_label(tag):
    cities = ["Houston", "Webster"]
    content = str(tag.contents)
    for city in cities:
        if tag.name == "label" and city in content:
            return True
    return False

def crawl_aegis_aerospace():
    # Aegis Aerospace
    url = "https://recruitingbypaycor.com/career/CareerHome.action?clientId=8a7883c68e824f98018ec4d2919c1832"
    web = Webber(url, wait_time=2, headless=True)
    
    aegis_label_xpath = "//label[@onclick]"
    
    web.click_xpath(aegis_label_xpath)
    
    # Get list of local postings
    labels = web.soup.find_all(houston_or_webster_label)
    posting_list = []
    for label in labels:
        posting_list.append(label.parent)
    
    # Get list of individual postings
    job_list = []
    for pl in posting_list:
        jobs = pl.find_all("li", class_="gnewtonNode")
        job_list.extend(jobs)
    
    # Extract job names and links
    for job in job_list:
        first_link = job.find("a")
        web.add_a_to_dict(web.output_dict, first_link)
    
    output_file = "company_jobs/Aegis_Aerospace.json"
    web.output_dict_to_json(output_file)

def crawl_aerospace_corp():
    # Goes through the aerospace corp jobs from a GET request
    # returns a dictionary of names associated with urls
    search_url = "https://talent.aerospace.org/api/apply/v2/jobs"

    dict_length = 0
    web = Webber(search_url, implementation="requests")
    
    #Aerospace corp only gives 10 positions at a time, so it has to loop
    for num in range(0, 100, 10):
        request_data = {
            "domain": "aerospace.org",
            "start": 0 + num,
            "num": 10 + num,
            "exclude_pid": 790299282440,
            "location": "Houston%2C%20TX",
            "pid": 790299282440,
            "domain": "aerospace.org",
            "sort_by": "relevance"
        }
        web.update_request(new_request_data=request_data)
        
        position_list = json.loads(web.request.text)["positions"]
        for position in position_list:
            web.output_dict[position["name"]] = position["canonicalPositionUrl"]
        
        # stop the for loop if we're not getting anything new
        if dict_length == len(web.output_dict):
            break
        else:
            dict_length = len(web.output_dict)
    
    output_file = "company_jobs/Aerospace_Corporation.json"
    web.output_dict_to_json(output_file)

def crawl_amentum():
    # Goes through Jacobs/Amentum's page with GET requests
    
    amentum_url = "https://www.amentumcareers.com/jobs/search"
    web = Webber(amentum_url, implementation="requests")
    #am_job_dict = {}
    dict_length = 0
    for num in range(1, 10):
        request_data = {
            "block_uid": "f3e697d2dea198aaa8c83cdf4b3e741f",
            "block_index": 0,
            "page_row_uid": "bd2bd9f0fc912c493d727d88a2fe5227",
            "page_row_index": 1,
            "page_version_uid": "dc558f4cb26594b844a0dbe98be37425",
            "page": num,
            #"location_uids": ,
            #"sort": ,
            #"search_workplace_types": ,
            #"search_employment_types": ,
            "employment_type_uids%5B%5D": "c32bed040f60654c22f167dfe5e10d09",
            #"search_country_codes": ,
            #"search_states": ,
            #"search_cities": ,
            "cities[]": "Houston",
            #"query": ,
        }
        web.update_request(new_request_data=request_data)
        print(web.request)
        #r = requests.get(amentum_url, request_data)
        #soup = BeautifulSoup(r.text, "html.parser")
        job_rows = web.soup.find_all("td", class_="job-search-results-title")
        for job in job_rows:
            link = job.find("a")
            Webber.add_a_to_dict(web.output_dict, link)
        
        if dict_length == len(web.output_dict):
            break
        else:
            dict_length = len(web.output_dict)
        time.sleep(4)
    output_file = "company_jobs/Amentum.json"
    web.output_dict_to_json(output_file)

def crawl_intuitive():
    
    url = "https://workforcenow.adp.com/mascsr/default/mdf/recruitment/recruitment.html?cid=b0e24f83-6e4d-492d-9d6a-bc0fea197d6a&ccId=19000101_000001&lang=en_US"
    web = Webber(url, wait_time=2, headless=False)
    
    
    view_all_xpath = "//*[@id=\"recruitment_careerCenter_showAllJobs\"]"
    web.click_xpath(view_all_xpath)
    for i in range (1, 3):
        web.scroll_to_bottom()
    web.refresh_soup()
    
    opening_list = web.soup.find_all("div", class_="current-openings-details")
    for d in opening_list:
        title = d.find("span", class_="current-opening-title")
        location = d.find("label", class_="current-opening-location-item")
        if "Houston" in location.string:
            key = web.strip_all_white_space(title.string)
            web.output_dict[key] = url
    
    output_file = "company_jobs/Intuitive_Machines.json"
    web.output_dict_to_json(output_file)

def crawl_bastion():
    
    bastion_url = "https://bastiontechnologies.applicantpro.com/jobs/?state=TX"
    
    web = Webber(bastion_url, wait_time=3, headless=True)
    
    #xpath in words: find all links under a div with the following class
    xpath = "//div[@class=\"job-name col-auto text-primary pl-0 pr-0\"]/a[@class=\"listing-url\"]"
    job_links = web.search_xpath(xpath)
    
    for link in job_links:
        key = link.text
        link_href = link.get_attribute("href")
        web.output_dict[key] = link_href
    
    output_file = "company_jobs/Bastion_Technologies.json"
    web.output_dict_to_json(output_file)

def crawl_mri():
    
    mri_url = "https://mricompany.applicantpro.com/jobs/?ref=mricompany.com"
    
    web = Webber(mri_url, wait_time=3, headless=True)
    
    #xpath in words: find all links under a div with the following class
    xpath = "//div[@class=\"job-name col-auto text-primary pl-0 pr-0\"]/a[@class=\"listing-url\"]"
    job_links = web.search_xpath(xpath)
    
    for link in job_links:
        key = link.text
        link_href = link.get_attribute("href")
        web.output_dict[key] = link_href
    
    output_file = "company_jobs/MRI_Technologies.json"
    web.output_dict_to_json(output_file)

def crawl_aerodyne():
    #aerodyne_url = "https://aerodyneindustries.hua.hrsmart.com/hr/ats/JobSearch/search"
    
    aerodyne_parent_url = "https://aerodyneindustries.hua.hrsmart.com"
    search_url = aerodyne_parent_url + "/hr/ats/JobSearch/search"
    
    request_data = {
        "submittedFormId": "jobSearchForm",
        "with_all": "",
        "with_at_least": "",
        "with_exact": "",
        "without": "",
        "location[]": "city:Houston|country:1|state:44",
        "zip_code_radius": "",
        "zip_code": "",
        "hua_country_id": 1,
        "ats_requisition_code": "",
        "search_jobs": "Search",
    }
    web = Webber(search_url, implementation="requests", 
                 request_data=request_data)
    
    span_elements = web.soup.find("tbody").find_all("span")
    for span in span_elements:
        job_url = aerodyne_parent_url + span.parent.get("href")
        web.output_dict[span.string] = job_url
    
    output_file = "company_jobs/Aerodyne_Industries.json"
    web.output_dict_to_json(output_file)

def crawl_barrios():
    search_url = "https://careers-barrios.icims.com/jobs/search?ss=1&searchLocation=12781-12827-Houston"
    
    request_data = {
        "ss": 1,
        "searchLocation": "12781-12827-Houston",
        "in_iframe": 1,
    }
    
    web = Webber(search_url, implementation="requests", 
                 request_data=request_data)
    
    link_elements = web.soup.find_all("a", class_="iCIMS_Anchor")
    for le in link_elements:
        job_name = web.strip_all_white_space(le.find("h3").string)
        web.output_dict[job_name] = le.get("href")
    
    output_file = "company_jobs/Barrios_Technology.json"
    web.output_dict_to_json(output_file)

def crawl_paragon():
    search_url = "https://paragonsdc.applicantpro.com/jobs/"
    
    web = Webber(search_url, wait_time=3, headless=True)
    time.sleep(2)
    web.refresh_soup()
    link_elements = web.soup.find_all("a", class_="listing-url")
    for le in link_elements:
        web.add_a_to_dict(web.output_dict, le)
    
    output_file = "company_jobs/Paragon.json"
    web.output_dict_to_json(output_file)

def crawl_axiom():
    site_url = "https://axiomspace.wd5.myworkdayjobs.com/External_Career_Site"
    search_url = "https://axiomspace.wd5.myworkdayjobs.com/wday/cxs/axiomspace/External_Career_Site/jobs"
    request_data = {
        "appliedFacets": {},
        "limit": 20,
        "offset": 0,
        "searchText": ""
    }
    
    web = Webber(search_url, request_data=request_data, implementation="post")
    position_list = json.loads(web.request.text)["jobPostings"]
    
    for pos in position_list:
        web.output_dict[pos["title"]] = site_url + pos["externalPath"]
    
    output_file = "company_jobs/Axiom.json"
    web.output_dict_to_json(output_file)
'''
def crawl_voyager():
    
    search_url = "https://voyagerspace.com/wp-json/adp/jobs-rss/"
    web = Webber(search_url, implementation="rss")
    print(web.rss[0])
'''
#crawl_aegis_aerospace()
#crawl_aerodyne()
#crawl_mri()
#crawl_bastion()
#crawl_intuitive()
#crawl_aerospace_corp()
#crawl_barrios()
#crawl_paragon()
#crawl_axiom()
crawl_amentum()


#Webber.write_to_json("MRI_Technologies.json", crawl_mri())
#Webber.write_to_json("Aegis_Aerospace.json", crawl_aegis_aerospace())
#Webber.write_to_json("Aerospace_Corporation.json", crawl_aerospace_corp())
#Webber.write_to_json("Amentum.json", crawl_amentum())
#Webber.write_to_json("Intuitive_Machines.json", crawl_intuitive())
#Webber.write_to_json("Bastion_Technologies.json", crawl_bastion())
#Webber.write_to_json("Aerodyne_Industries.json", crawl_aerodyne())