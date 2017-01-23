# coding: utf-8

import pyexcel as p
import os


class XlsDuplicateParser():

    def __init__(self):
        self.sheets_to_compare = []

    def add_sheet_to_compare(self, file_name):
        self.sheets_to_compare.append(file_name)

    def browse_rows(self):
        set1 = set(elt for row in self.sheet1 for elt in row)
        set2 = set(elt for row in self.sheet2 for elt in row)

        result = set1 & set2 - {'', None}
        return result

    def browse_rows_multiple_files(self, files):
        sets = []
    
        for f in files:
            file_name, file_extension = os.path.splitext(f)
            library = "pyexcel-{}".format(file_extension[1:])
            
            sheet = p.get_sheet(file_name=f, library=library)
            sets.append(set(elt for row in sheet for elt in row))
    
        result = set.intersection(*sets) - {'', None}
        return result