# coding: utf-8

import pyexcel as p


class XlsDuplicateParser():

    def __init__(self, file_name1, file_name2):
        self.sheet1 = p.get_array(file_name=file_name1)
        self.sheet2 = p.get_array(file_name=file_name2)
        self.browse_spreadsheet_rows()

    def check_if_cells_match(self, cell1, cell2):
        """ Return True if 2 cells have the same content (not type) """
        if (cell1 == cell2) and (cell1 != '') and (cell2 != ''):
            return True

    def browse_spreadsheet_rows(self):
        """ Loop trough 2 worksheets to compare all cells """
        cells_checked_count = 0
        matches = []

        for row_first_sheet in self.sheet1:
            for cell_first_sheet in row_first_sheet:
                cells_checked_count += 1
                if cells_checked_count % 250 == 0:
                    print('({} cells checked)'.format(cells_checked_count))
                for row_second_sheet in self.sheet2:
                    for cell_second_sheet in row_second_sheet:
                        if self.check_if_cells_match(
                                cell_first_sheet, cell_second_sheet):
                            if cell_first_sheet not in matches:
                                print(
                                    'MATCH : {} ({} cells checked)'.format(
                                        cell_first_sheet,
                                        cells_checked_count
                                    )
                                )
                                matches.append(cell_first_sheet)

        print('FINISHED : {} cells checked'.format(cells_checked_count))


file_name1 = input('Entrez le nom du 1er fichier :')
file_name2 = input('Entrez le nom du 2ème fichier :')

t = XlsDuplicateParser(file_name1, file_name2)
