# coding: utf-8

import pyexcel as p
import phonenumbers as pn
import types
import datetime
import os

class PnParser():

    def __init__(self):
        self.sheets_to_parse    = []
        self.sheets_parsed      = []

    def add_sheet_to_parse(self, sheet_name):
        self.sheets_to_parse.append(sheet_name)

    def parse(self):
        for sheet in self.sheets_to_parse:
            self.parse_file(sheet)

    def remove_swisscom_shit(self, cell):
        if cell[0:3] == '120':
            cell = cell[3:]
        elif cell[0:2] == '12':
            cell = cell[2:]
        return cell

    def parse_file(self, sheet_name):
        file_name, file_extension = os.path.splitext(sheet_name)
        library = "pyexcel-{}".format(file_extension[1:])

        sheet = p.get_sheet(file_name=sheet_name)
        numbers_modified = []

        for i, row in enumerate(sheet):
            for c, cell in enumerate(row):
                if isinstance(cell, str) or isinstance(cell, int):
                    cell = self.remove_swisscom_shit(str(cell))
                    if pn.is_possible_number_string(str(cell), 'CH'):
                        raw_number = pn.parse(str(cell), 'CH')
                        if pn.is_valid_number(raw_number):
                            formated_number = pn.format_number(raw_number, pn.PhoneNumberFormat.E164)[1:]
                            sheet.cell_value(i, c, new_value=formated_number)
                            numbers_modified.append({'old': cell, 'new': formated_number})
                elif isinstance(cell, datetime.date):
                    d = cell.strftime(u'%d/%m/%Y %H:%M:%S')
                    sheet.cell_value(i, c, new_value=d)

        a = sheet_name.split('.')
        ext = a.pop()
        new_sheet_name = ''.join(a)
        new_sheet_name = '{}_parsed.{}'.format(new_sheet_name, ext)

        sheet.save_as(new_sheet_name)
        self.sheets_parsed.append(
            {
                'name': new_sheet_name,
                'numbers_modified': numbers_modified
            }
        )
        sheet = None

    def browse(self, files):
        sets = []
    
        for f in files:
            file_name, file_extension = os.path.splitext(f)
            library = "pyexcel-{}".format(file_extension[1:])
            
            sheet = p.get_sheet(file_name=f, library=library)
            sets.append(set(elt for row in sheet for elt in row if pn.is_possible_number_string(str(elt), 'CH')))
    
        result = set.intersection(*sets) - {'', None}
        return result