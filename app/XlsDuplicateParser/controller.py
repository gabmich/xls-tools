# coding: utf-8

import pyexcel as p


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
            sheet = p.get_array(file_name=f)
            sets.append(set(elt for row in sheet for elt in row))
    
        result = set.intersection(*sets) - {'', None}
        return result