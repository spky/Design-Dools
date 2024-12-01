import pandas as pd
import os
from bs4 import BeautifulSoup

class WebsiteTable:
    """A class used to read through and set up arbitrary table set ups, mainly on 
    McMaster-Carr's website"""
    
    def __init__(self, file_path):
        
        self.absolute_path = os.path.abspath(file_path)
        self.file_name = os.path.basename(file_path)
        self.folder = os.path.dirname(self.absolute_path)
        self.titles = []
        self.active_table = None
        with open(self.absolute_path, "r") as file:
            self.soup = BeautifulSoup(file, "html.parser")
        
        self.tables = self.soup.find_all("table")
    
    def print_table_info(self):
        for count, t in enumerate(self.tables):
            print("Table " + str(count))
            self.print_attributes(t, "    ")
    
    def activate_table(self, table_number):
        self.active_table = self.tables[table_number]
        self.header_element = self.active_table.find("thead")
        self.titles = self.span_child_text(self.header_element)
    
    @staticmethod
    def span_child_text(element):
        # Returns the text inside of all of the element's child span elements in a 
        # list format
        spans = element.find_all("span")
        texts = []
        for s in spans:
            texts.append(s.text)
        return texts
    
    @staticmethod
    def print_attributes(soup, indent="    "):
        if not soup.attrs:
            print(indent + "No attributes")
        else:
            for key in soup.attrs:
                print(indent + key + ": " + str(soup[key]))
    
    @staticmethod
    def print_children_tags(soup, indent="", print_list=[]):
        i_indent = "    "
        print(indent + soup.name)
        
        for count, child in enumerate(soup.children):
            print(indent + i_indent + "element " + str(count) + ": " + child.name)
            if "attributes" in print_list:
                print(indent + i_indent + "Attributes:")
                WebsiteTable.print_attributes(child, i_indent*2)
            if "text" in print_list:
                print(indent + i_indent + "string: " + str(child.string))
            if "children" in print_list:
                print(indent + i_indent + "Children: ")
                WebsiteTable.print_children_tags(child, i_indent*2)