import datetime
import re
import xlrd


def text_to_num_format(cell):
    if cell.ctype == xlrd.XL_CELL_TEXT:
        digits = re.findall(r'\d+', cell.value)
        cell_value = int(digits[0]) if digits != [] else None
        return cell_value


def text_num_to_int(cell):
    if cell.ctype == xlrd.XL_CELL_TEXT:
        digits = re.findall(r'\d+', cell.value)
        cell_value = int(digits[0]) if digits != [] else None
        return cell_value
    elif cell.ctype == xlrd.XL_CELL_NUMBER and int(cell.value) == cell.value:
        cell_value = int(cell.value)
        return cell_value
    else:
        cell_value = None
        return cell_value


def date_format(cell, wb):
    if cell.ctype == xlrd.XL_CELL_DATE:
        cell_value = datetime.datetime(*xlrd.xldate_as_tuple(cell.value, wb.datemode))
        return cell_value
    else:
        cell_value = None
        return cell_value


def int_format(cell):
    if cell.ctype == xlrd.XL_CELL_NUMBER and int(cell.value) == cell.value:
        cell_value = int(cell.value)
        return cell_value


def text_format(cell):
    if cell.ctype == xlrd.XL_CELL_TEXT:
        cell_value = cell.value
        cell_value = cell_value.strip()
        return cell_value