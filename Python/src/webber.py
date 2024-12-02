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