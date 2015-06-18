# XXX Fill out docstring!
"""

"""
from __future__ import print_function
import logging
import sys

from errors import DataError
log = logging.getLogger(__name__)
__author__ = 'wgibb'


# http://code.activestate.com/recipes/577058/
def query_yes_no(question, default="yes"):
    """Ask a yes/no question via raw_input() and return their answer.
    "question" is a string that is presented to the user.
    "default" is the presumed answer if the user just hits <Enter>.
        It must be "yes" (the default), "no" or None (meaning
        an answer is required of the user).

    The "answer" return value is one of "yes" or "no".
    """
    valid = {"yes": True, "y": True, "ye": True,
             "no": False, "n": False}
    if not default:
        prompt = " [y/n] "
    elif default == "yes":
        prompt = " [Y/n] "
    elif default == "no":
        prompt = " [y/N] "
    else:
        raise ValueError("invalid default answer: '%s'" % default)

    while True:
        sys.stdout.write(question + prompt)
        # XXX not python3 friendly.
        choice = raw_input().lower()
        if default is not None and choice == '':
            return valid[default]
        elif choice in valid:
            return valid[choice]
        else:
            sys.stdout.write("Please respond with 'yes' or 'no' " "(or 'y' or 'n').\n")


def print_items_keys(iterable):
    for thing in iterable:
        for attr, value in thing.__dict__.items():
            print(attr, value)
        print('================================')


''' general helper function for ws row/column access'''


def get_index_by_value(cell_list, value):
    """ given a list of cells and a value, return the index in the list of
    that cell value.  this returns the index of the first time the value
    is found.

    if the value is not found, return false
    """
    index = False
    for cell in cell_list:
        if cell.value == value:
            try:
                index = cell_list.index(cell)
            except ValueError:
                log.exception('Failed to get index for value after finding a match')
                return index
            break
    return index


def stats(wb, sheetlist, value, verbose=False):
    """generate counts of different values in a given column, based on column name
    this is done across a workbook"""
    dicty = {}
    for sheet in sheetlist:
        ws = wb.get_sheet_by_name(sheet)
        row1 = ws.rows[0]
        index = get_index_by_value(row1, value)
        for cell in ws.columns[index]:
            if cell.value in dicty:
                dicty[cell.value] += 1
            else:
                dicty[cell.value] = 1
                if verbose:
                    log.info('Found new cell.value in sheet: {}'.format(sheet))
    return dicty


def build_data_from_fields(ws, fields):
    header = ws.rows[0]
    header_index = {}
    ret = []
    for key in fields.required_attributes:
        index = get_index_by_value(header, key)
        if index is False:
            raise DataError('Failed to obtain header value: {}'.format(key))
        header_index[key] = index
    ws_end = len(ws.columns[0]) - 1
    if not ws_end:
        raise DataError('Failed to find end of the ws data')
    # This skips the header row, as we are embedding all of the required data into dictionaries.
    for row in ws.rows[1:ws_end]:
        d = {key: row[index].value for key, index in header_index.items()}
        ret.append(d)
    return ret