from __future__ import print_function

'''
analyzer.py

library for handling root data from winrhizotron

this whole file should really be broken between a library file and a program

'''
import argparse
import logging
import os
import sys

import openpyxl


# logging config
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s: %(levelname)s: %(message)s [%(filename)s:%(funcName)s]')
log = logging.getLogger(__name__)


def get_test_ws():
    wb = openpyxl.load_workbook(filename=r'unedited.xlsx')
    ws = wb.get_sheet_by_name('T008 Root Synth')
    return ws


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


''' helper functions for finding the import values from the root data table'''


def build_root_data_table(ws, rootindexdata):
    """ BUILD A LIST OF DATA REPRESENTING EACH ROW IN WORKSHEET UP
    TO THE SYNTHESIS TABLE
    """
    rootdata = []
    # get index values for each important column
    tableheader = ws.rows[0]
    for key in rootindexdata:
        index = get_index_by_value(tableheader, key)
        if not index:
            log.error('Failed to obtain header value: {}'.format(key))
            return False
        rootindexdata[key] = index
    endof_rootdata = get_index_at_null(ws.columns[0])
    if not endof_rootdata:
        logging.error('Failed to find end of root data.')
        return False
    # this slice includes the header row in rootData.  This lets us
    # rebuild indices, since the dict iterator is not sorted
    for row in ws.rows[:endof_rootdata]:
        rootrow = []
        for key in rootindexdata:
            value = row[rootindexdata[key]].value
            rootrow.append(value)
        rootdata.append(rootrow)
    return rootdata


def process_worksheet(ws, exceptions_list=False):
    """Process a worksheet, and build a tube object containing all the root data.
    This data can then be dumped out when all tubes have been processed."""
    root_index_data = {
        'RootName': 0,
        'Location#': 0,
        'Session#': 0,
        'BirthSession': 0,
        'TotAvgDiam(mm/10)': 0,
        'TipLivStatus': 0,
        'RootNotes': 0,
        'NumberOfTips': 0,
        'Date': 0,
    }
    tube_number = get_tube_number(ws)
    # exception handling
    exceptions = None
    if exceptions_list:
        try:
            exceptions = exceptions_list[tube_number]
        except KeyError:
            log.exception('No exceptions list found for tube number: %s' % (str(tube_number)))
    # basic information
    tube = Tube(tube_number)
    logging.debug('Processing Data for tube #: %s' % (str(tube_number)))
    #
    root_data = build_root_data_table(ws, root_index_data)
    if not root_data:
        log.error('Failed to get root data from sheet: {}'.format(ws.title))
        return False
    header_row = root_data[0]
    for key in root_index_data:
        try:
            index = header_row.index(key)
        except ValueError:
            log.exception('Failed to obtain header value: {}'.format(key))
            return False
        root_index_data[key] = index
    # this loop will not handle exceptional roots, it just builts a list of root 
    # object for each row of root data
    logging.debug('Building roots from root_data')
    max_session_count = 1
    session_dates = {}
    roots = []
    for row in root_data[1:]:
        if exceptions:
            # noinspection PyUnusedLocal
            for i in exceptions:
                # noinspection PyPep8Naming,PyUnusedLocal
                rootName = row[root_index_data['RootName']]
                # noinspection PyPep8Naming,PyUnusedLocal
                rootLocation = row[root_index_data['Location#']]
                # noinspection PyPep8Naming,PyUnusedLocal
                rootSession = row[root_index_data['Sesssion#']]
                #
                # EXCEPTION ROOT HANDLING WOULD WOULD GO HERE
                #
                pass
        root = root_from_row(row, root_index_data)
        roots.append(root)
        # process session statistics
        if root.session > max_session_count:
            max_session_count = root.session
            logging.debug('MaxSessionCount updated to %s' % (str(max_session_count)))
        if root.session not in session_dates:
            session_dates[root.session] = row[root_index_data['Date']]
            logging.debug('Inserted session %s- Date %s ' % (str(root.session), session_dates[root.session],))
    # no longer need to keep this data in memory
    del root_data
    # need to know what session value to use for finalizing root values.
    tube.maxSessionCount = max_session_count
    tube.sessionDates = session_dates
    # add each root to the tube
    logging.debug('Inserting roots')
    final_roots = []
    for root in roots:
        tube.insert_or_update_root(root)
        if root.session == max_session_count:
            final_roots.append(root)
    # finalize roots
    logging.debug('Finalizing roots')
    for root in final_roots:
        status = tube.finalize_root(root)
        if not status:
            logging.error('Failed to finalize root %s' % (str(root.identity)))
            log.error(root.identity)

    # need to calculate alive/dead tip numbers
    tstats = tip_stats(tube)
    for r in tube.roots:
        r.aliveTipsAtBirth = tstats[r.tipIdentity]
        if r.goneSession:
            gone_identity = r.goneSession, r.location
            r.aliveTipsAtGone = tstats[gone_identity]
    tube.tipStats = tstats

    logging.debug('Identified %s roots in sheet: %s' % (str(len(roots)), ws.title))
    logging.debug('Found %s roots in tube' % (str(len(tube.roots))))
    return tube


# tip stats code

def calculate_alive_tip_stats(tube):
    alive_tips = {}
    for existingRoot in tube.roots:
        tip_id = existingRoot.tipIdentity
        if tip_id not in alive_tips:
            alive_tips[tip_id] = 1
        else:
            alive_tips[tip_id] += 1
    return alive_tips


def calculate_gone_tip_stats(din, tube):
    d = {}
    for i, c in sorted(din.items()):
        for r in tube.roots:
            if (r.tipIdentity == i) and r.goneSession:
                gi = r.goneSession, r.location
                if gi not in d:
                    d[gi] = -1
                else:
                    d[gi] -= 1
    return d


def tip_stats(tube):
    # get count of alive tips and the locations/sessions where roots die
    alive_stats = calculate_alive_tip_stats(tube)
    gone_stats = calculate_gone_tip_stats(alive_stats, tube)
    # build a list of all root locations
    locations = []
    for r in tube.roots:
        if r.location not in locations:
            locations.append(r.location)
    # build a mapping of sessions to lists of locations
    loc_sessions = {}
    for L in locations:
        for k in alive_stats:
            sess, loc = k
            if L == loc:
                if L not in loc_sessions:
                    loc_sessions[L] = [sess]
                else:
                    loc_sessions[L].append(sess)
    for L in locations:
        for k in gone_stats:
            sess, loc = k
            if L == loc:
                if L not in loc_sessions:
                    loc_sessions[L] = [sess]
                else:
                    loc_sessions[L].append(sess)
    # unique & sort the lists in loc_sessions
    for k in loc_sessions:
        i = list(set(loc_sessions[k]))
        i.sort()
        loc_sessions[k] = i
    # sum the alive and dead roots over time
    total_stats = {}
    for k in loc_sessions:
        temp = 0
        for sess in loc_sessions[k]:
            key = sess, k
            if key in alive_stats:
                temp = alive_stats[key] + temp
            if key in gone_stats:
                temp = gone_stats[key] + temp
            total_stats[key] = temp
    # return totalStats
    return total_stats


# root processing code

def root_from_row(row, indexdict):
    """ build a root object from row, using indexdict for mappings"""
    # identification information
    root_name = row[indexdict['RootName']]
    location = row[indexdict['Location#']]
    birth_session = row[indexdict['BirthSession']]
    root = Root(root_name, location, birth_session)
    # check to see if anomalous root
    if row[indexdict['NumberOfTips']] == 1:
        root.anomaly = False
        if row[indexdict['TipLivStatus']].startswith('A'):
            root.isAlive = 'A'
        if row[indexdict['TipLivStatus']].startswith('G'):
            root.isAlive = 'G'
    else:
        root.anomaly = True
        root.isAlive = 'A'
    # root metadata
    root.session = row[indexdict['Session#']]
    if root.isAlive.startswith('G'):
        root.goneSession = root.session
    # order & diameter data
    root.order = row[indexdict['RootNotes']]
    root.avgDiameter = row[indexdict['TotAvgDiam(mm/10)']]
    return root


def get_tube_number(ws):
    """ get tube number from a given worksheet.  this assumes that """
    row1 = ws.rows[0]
    tube_index = get_index_by_value(row1, 'Tube#')
    tube_number = ws.columns[tube_index][1].value
    if tube_number:
        return tube_number
    else:
        return False


''' helper function for working with openpyxl'''


def get_root_sheets(wb):
    """given a workbook, return all sheet names that end in "Root Synth"\""""
    root_sheets = []
    sheets = wb.get_sheet_names()
    for sheet in sheets:
        if sheet.endswith('Root Synth'):
            root_sheets.append(sheet)
    return root_sheets


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


def print_items_keys(iterable):
    for thing in iterable:
        for attr, value in thing.__dict__.items():
            print(attr, value)
        print('================================')


# main function support
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


def write_out_data(output, tubes):
    """
    write out the data from tubes in a xlsx workbook

    header data chosen by MSS
    """
    # setup header indices
    header = ['RootName', 'Tube#', 'Location#', 'BirthSession', 'Birth Date',
              'Birth Year', 'GoneSession', 'Gone Date', 'Gone Year', 'Censored',
              'AliveTipsAtBirth', 'AliveTipsAtGone', 'Avg Diameter (mm/10)',
              'Order', 'Highest Order']
    header_index = {}
    # header index specifically for use with get_column_letter()
    for i in header:
        header_index[i] = header.index(i) + 1
    headertorootmapping = {
        'RootName': 'rootName',
        'Location#': 'location',
        'BirthSession': 'birthSession',
        'GoneSession': 'goneSession',
        'Censored': 'censored',
        'AliveTipsAtBirth': 'aliveTipsAtBirth',
        'AliveTipsAtGone': 'aliveTipsAtGone',
        'Avg Diameter (mm/10)': 'avgDiameter',
        'Order': 'order',
        'Highest Order': 'highestOrder',
    }

    # setup workbook
    wb = openpyxl.Workbook()
    ws = wb.worksheets[0]
    ws.title = 'Compiled Data'
    row_index = 1
    # write out header row
    for i in header:
        col_indx = header_index[i]
        col = openpyxl.cell.get_column_letter(col_indx)
        ws.cell('%s%s' % (col, row_index)).value = i
    row_index += 1
    # process the root data from each tube
    for tube in tubes:
        tube_number = tube.tubeNumber
        logging.debug('Writing out data for tube#: %s' % (str(tube_number)))
        for root in tube.roots:
            # write header
            col_indx = header_index['Tube#']
            col = openpyxl.cell.get_column_letter(col_indx)
            ws.cell('%s%s' % (col, row_index)).value = tube_number
            # build a dictionary of all the root items
            root_dict = root.__dict__
            for value in header:
                if value in headertorootmapping:
                    root_value = root_dict[headertorootmapping[value]]
                    col_indx = header_index[value]
                    col = openpyxl.cell.get_column_letter(col_indx)
                    ws.cell('%s%s' % (col, row_index)).value = root_value
                #
                # These two elif statements are not dynamic
                #
                elif value == 'Birth Date':
                    birth_session = root_dict['birthSession']
                    birth_date = tube.sessionDates[birth_session]
                    # specific to the data provided by MSS
                    birth_year = birth_date.split('.')[0]
                    col_indx = header_index[value]
                    col = openpyxl.cell.get_column_letter(col_indx)
                    ws.cell('%s%s' % (col, row_index)).value = birth_date
                    col_indx = header_index['Birth Year']
                    col = openpyxl.cell.get_column_letter(col_indx)
                    ws.cell('%s%s' % (col, row_index)).value = birth_year
                elif value == 'Gone Date':
                    # make sure root has kicked the bucket first.
                    if root_dict['censored'] == 0:
                        gone_session = root_dict['goneSession']
                        gonedate = tube.sessionDates[gone_session]
                        # specific to data provided by MSS
                        goneyear = gonedate.split('.')[0]
                        col_indx = header_index[value]
                        col = openpyxl.cell.get_column_letter(col_indx)
                        ws.cell('%s%s' % (col, row_index)).value = gonedate
                        col_indx = header_index['Gone Year']
                        col = openpyxl.cell.get_column_letter(col_indx)
                        ws.cell('%s%s' % (col, row_index)).value = goneyear
                elif value not in ['Birth Year', 'Gone Year', 'Tube#']:
                    log.warning('Unhandled value in header: {}'.format(value))
            row_index += 1
    # save results
    # noinspection PyBroadException
    try:
        wb.save(filename=output)
    except Exception:
        log.exception('General error handler')
        return False
    return True


#
# Main program functions
#


def main(options):
    #
    #   DOES NOT HANDLE EXCEPTION LISTS
    #

    # file path verification and options handling
    if not options.verbose:
        logging.info('Output will not be verbose')
        logging.disable(logging.DEBUG)

    if not os.path.isfile(options.src_file):
        logging.error('specified source is not a file')
        sys.exit(-1)

    if os.path.exists(options.output):
        logging.warning('Specified output file already exists.\n')
        if not query_yes_no('Do you want to overwrite that file?', 'no'):
            logging.info('Exiting')
            sys.exit(-1)

    # open up src file
    # noinspection PyBroadException
    try:
        wb = openpyxl.load_workbook(filename=options.src_file)
    except Exception:
        log.exception('general exception handler')
        sys.exit(-1)

    # get root data
    sheets = get_root_sheets(wb)
    if not sheets:
        logging.error('Could not find any WinRhizotron data sheets in the src file')
        logging.error('Make sure your workbook worksheets end in the string "Root Synth"')
        sys.exit(-1)

    tubes = []
    for sheet in sheets:
        # noinspection PyBroadException
        try:
            ws = wb.get_sheet_by_name(sheet)
        except Exception:
            log.exception('Error occured attempting to get worksheet: {}'.format(sheet))
            # attempt to get the next worksheet
            continue
        tube = process_worksheet(ws)
        if not tube:
            log.error('Was not able to process the tube data from worksheet: {}'.format(sheet))
        tubes.append(tube)
    if not tubes:
        logging.error('Was unable to get any tube data')
        sys.exit(-1)

    # write out the data for each tube into the new worksheet
    foo = write_out_data(options.output, tubes)
    if foo:
        log.info('Sucesfully wrote out data to {}'.format(options.output))
    else:
        logging.info('Did not write data out.')
        sys.exit(-1)

    logging.info('Done processing all data')
    sys.exit(0)


def root_options():
    parser = argparse.ArgumentParser(prog=__name__)
    parser.add_argument('-s', '--source', dest='src_file', required=True, type=str, action='store',
                        help='Source xlsx file to process')
    parser.add_argument('-e', '--exceptions', dest='exceptions', type=str, action='store', default=None,
                        help='A xlsx file containing root rewrite data for shifted roots')
    parser.add_argument('-o', '--output', dest='output', required=True, type=str, action='store',
                        help='Define the output file (xlsx)')
    parser.add_argument('-v', '--verbose', dest='verbose', default=False, action='store_true',
                        help='Enable verbose output')
    return parser


if __name__ == "__main__":
    p = root_options()
    opts = p.parse_args()

    main(opts)
