import re
from itertools import zip_longest

from . import styles

# Apply functions apply a chosen structure to the row of the excel file the
# openpyxl generator is currently on. Every apply function has a match function
# that returns a boolean indicating if that apply function needs to be applied,
# based on the row of the docx table of content line.

regex_description = re.compile('\|\w*\|')
regex_numbers = re.compile('\d+')


def apply_row_with_sum(excel_row, docx_line, generator_through_sheet):
    """"""
    docx_line[-2:-2] = [''] * 4
    for excel_cell, docx_cell in zip_longest(excel_row, docx_line, fillvalue=''):
        excel_cell.value = docx_cell
        excel_cell.fill = styles.fill_grey
    index_row = excel_row[0].row
    excel_row[8].value = '=SUM(J{}:J{})'.format(index_row + 1, index_row + 3)
    for i in range(3):
        next(generator_through_sheet)


def apply_row_bold(excel_row, docx_line, generator_through_sheet):
    """"""
    for excel_cell, docx_cell in zip_longest(excel_row, docx_line, fillvalue=''):
        excel_cell.value = docx_cell
        excel_cell.font = styles.font_bold


def apply_row_default(excel_row, docx_line, generator_through_sheet):
    """"""
    for excel_cell, docx_cell in zip(excel_row, docx_line):
        excel_cell.value = docx_cell


# Match functions, based on the line of the docx-file table of content checks
# if the matching apply function needs to be applied.

def match_row_with_sum(docx_line):
    return re.search(regex_description, docx_line[-2])


def match_row_bold(docx_line):
    return (((int(re.findall(regex_numbers, docx_line[0])[-1]) % 10) == 0)
           | (len(re.findall(regex_numbers, docx_line[0])) == 1))


def match_row_default(docx_line):
    return True