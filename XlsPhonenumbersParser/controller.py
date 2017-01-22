# coding: utf-8

import pyexcel as p
import phonenumbers as pn
import types
import datetime

def date_to_format(value, target_format):
    """Convert date to specified format"""
    if target_format == str:
        if isinstance(value, datetime.date):
            ret = value.strftime("%d/%m/%y")
        elif isinstance(value, datetime.datetime):
            ret = value.strftime("%d/%m/%y")
        elif isinstance(value, datetime.time):
            ret = value.strftime("%H:%M:%S")
    else:
        ret = value
    return ret


class XlsPhonenumbersParser():

    def __init__(self):
        self.sheets = []
        self.new_sheets = []

    def add_file(self, file_name):
        self.sheets.append(file_name)

    def parse(self):
        for sheet in self.sheets:
            self.parse_file(sheet)

    def remove_swisscom_shit(self, cell):
        if cell[0:3] == '120':
            cell = cell[3:]
        elif cell[0:2] == '12':
            cell = cell[2:]
        return cell

    def parse_file(self, sheet_name):
        sheet = p.get_sheet(file_name=sheet_name)
        numbers_modified = []

        for i, row in enumerate(sheet):
            for c, cell in enumerate(row):
                cell = self.remove_swisscom_shit(str(cell))
                if pn.is_possible_number_string(str(cell), 'CH'):
                    raw_number = pn.parse(str(cell), 'CH')
                    if pn.is_valid_number(raw_number):
                        formated_number = pn.format_number(raw_number, pn.PhoneNumberFormat.E164)[1:]
                        sheet.cell_value(i, c, new_value=formated_number)
                        numbers_modified.append({'old': cell, 'new': formated_number})

        a = sheet_name.split('.')
        ext = a.pop()
        new_sheet_name = ''.join(a)
        new_sheet_name = '{}_parsed.{}'.format(new_sheet_name, ext)

        sheet.save_as(new_sheet_name)
        self.new_sheets.append(
            {
                'name': new_sheet_name,
                'numbers_modified': numbers_modified
            }
        )
