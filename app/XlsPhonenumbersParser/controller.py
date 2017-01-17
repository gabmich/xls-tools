# coding: utf-8

import pyexcel as p
import phonenumbers as pn


class XlsPhonenumbersParser():

    def load_file(self, file_name):
        self.file_name = file_name
        self.sheet = p.get_sheet(file_name=file_name)

    def parse_file(self):
        for i, row in enumerate(self.sheet):
            for c, cell in enumerate(row):
                if pn.is_possible_number_string(cell, "CH"):
                    raw_number = pn.parse(cell, 'CH')
                    formated_number = pn.format_number(raw_number, pn.PhoneNumberFormat.E164)
                    self.sheet.cell_value(i, c, new_value=formated_number)
        self.sheet.save_as('new_file.xlsx')

x = XlsPhonenumbersParser()
file_name = 'a1.xlsx'
x.load_file(file_name)
x.parse_file()
