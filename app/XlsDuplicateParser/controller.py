# coding: utf-8

import pyexcel as p


class XlsDuplicateParser():

    def set_sheet(self, file_name):
        return p.get_array(file_name=file_name)

    def set_first_sheet(self, file_name):
        """ Set the first file to compare """
        self.sheet1 = p.get_array(file_name=file_name)

    def set_second_sheet(self, file_name):
        """ Set the second file to compare """
        self.sheet2 = p.get_array(file_name=file_name)

    def check_if_cells_match(self, cell1, cell2):
        """ Return True if 2 cells have the same content (not type) """
        if (cell1 == cell2) and (cell1 != '') and (cell2 != ''):
            return True

    def browse_rows(self):
        set1 = set(elt for row in self.sheet1 for elt in row)
        set2 = set(elt for row in self.sheet2 for elt in row)

        result = set1 & set2 - {'', None}
        return result
