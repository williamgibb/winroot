# XXX Fill out docstring!
"""

"""
from __future__ import print_function
import logging
import sys

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



def get_index_at_null(cell_list):
    """ get the index value for the first NoneType object present in cell list"""
    index = False
    for cell in cell_list:
        if not cell.value:
            index = cell_list.index(cell)
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
