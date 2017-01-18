# coding: utf-8

import pyexcel as p
import phonenumbers as pn


class XlsPhonenumbersParser():

    def __init__(self):
        self.sheets = []

    def add_file(self, file_name):
        self.sheets.append(file_name)

    def parse(self):
        for sheet in self.sheets:
            self.parse_file(sheet)

    def parse_file(self, sheet_name):
        sheet = p.get_sheet(file_name=sheet_name)

        for i, row in enumerate(sheet):
            for c, cell in enumerate(row):
                cell = str(cell)
                if pn.is_possible_number_string(cell, "CH"):
                    raw_number = pn.parse(cell, 'CH')
                    formated_number = pn.format_number(raw_number, pn.PhoneNumberFormat.E164)
                    sheet.cell_value(i, c, new_value=formated_number)
        a = sheet_name.split('.')
        a.pop()
        new_sheet_name = ''.join(a)

        sheet.save_as('{}_parsed.xlsx'.format(new_sheet_name))

x = XlsPhonenumbersParser()
x.add_file('a1.xlsx')
x.add_file('fichier1.xlsx')
x.parse()
