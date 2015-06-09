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

import root
import tube
import utility


# logging config
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s: %(levelname)s: %(message)s [%(filename)s:%(funcName)s]')
log = logging.getLogger(__name__)


def get_test_ws():
    wb = openpyxl.load_workbook(filename=r'unedited.xlsx')
    ws = wb.get_sheet_by_name('T008 Root Synth')
    return ws


''' helper functions for finding the import values from the root data table'''


def build_root_data_table(ws, rootindexdata):
    """ BUILD A LIST OF DATA REPRESENTING EACH ROW IN WORKSHEET UP
    TO THE SYNTHESIS TABLE
    """
    rootdata = []
    # get index values for each important column
    tableheader = ws.rows[0]
    for key in rootindexdata:
        index = utility.get_index_by_value(tableheader, key)
        if not index:
            log.error('Failed to obtain header value: {}'.format(key))
            return False
        rootindexdata[key] = index
    endof_rootdata = utility.get_index_at_null(ws.columns[0])
    if not endof_rootdata:
        log.error('Failed to find end of root data.')
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


def process_worksheet(ws):
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
    # basic information
    t = tube.Tube(tube_number)
    log.debug('Processing Data for tube #: %s' % (str(tube_number)))
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
    log.debug('Building roots from root_data')
    max_session_count = 1
    session_dates = {}
    roots = []
    for row in root_data[1:]:
        root = root_from_row(row, root_index_data)
        roots.append(root)
        # process session statistics
        if root.session > max_session_count:
            max_session_count = root.session
            log.debug('MaxSessionCount updated to %s' % (str(max_session_count)))
        if root.session not in session_dates:
            session_dates[root.session] = row[root_index_data['Date']]
            log.debug('Inserted session %s- Date %s ' % (str(root.session), session_dates[root.session],))
    # no longer need to keep this data in memory
    del root_data
    # need to know what session value to use for finalizing root values.
    t.maxSessionCount = max_session_count
    t.sessionDates = session_dates
    # add each root to the tube
    log.debug('Inserting roots')
    final_roots = []
    for root in roots:
        t.insert_or_update_root(root)
        if root.session == max_session_count:
            final_roots.append(root)
    # finalize roots
    log.debug('Finalizing roots')
    for root in final_roots:
        status = t.finalize_root(root)
        if not status:
            log.error('Failed to finalize root %s' % (str(root.identity)))
            log.error(root.identity)

    # need to calculate alive/dead tip numbers
    tstats = tip_stats(t)
    for r in t.roots:
        r.aliveTipsAtBirth = tstats[r.tipIdentity]
        if r.goneSession:
            gone_identity = r.goneSession, r.location
            r.aliveTipsAtGone = tstats[gone_identity]
    t.tipStats = tstats

    log.debug('Identified %s roots in sheet: %s' % (str(len(roots)), ws.title))
    log.debug('Found %s roots in tube' % (str(len(t.roots))))
    return t


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
    r = root.Root(root_name, location, birth_session)
    # check to see if anomalous root
    if row[indexdict['NumberOfTips']] == 1:
        r.anomaly = False
        if row[indexdict['TipLivStatus']].startswith('A'):
            r.isAlive = 'A'
        if row[indexdict['TipLivStatus']].startswith('G'):
            r.isAlive = 'G'
    else:
        r.anomaly = True
        r.isAlive = 'A'
    # root metadata
    r.session = row[indexdict['Session#']]
    if r.isAlive.startswith('G'):
        r.goneSession = r.session
    # order & diameter data
    r.order = row[indexdict['RootNotes']]
    r.avgDiameter = row[indexdict['TotAvgDiam(mm/10)']]
    return r


def get_tube_number(ws):
    """ get tube number from a given worksheet.  this assumes that """
    row1 = ws.rows[0]
    tube_index = utility.get_index_by_value(row1, 'Tube#')
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



# main function support
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
        log.debug('Writing out data for tube#: %s' % (str(tube_number)))
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
        log.info('Output will not be verbose')
        logging.disable(logging.DEBUG)

    if not os.path.isfile(options.src_file):
        log.error('specified source is not a file')
        sys.exit(-1)

    if os.path.exists(options.output):
        log.warning('Specified output file already exists.\n')
        if not utility.query_yes_no('Do you want to overwrite that file?', 'no'):
            log.info('Exiting')
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
        log.error('Could not find any WinRhizotron data sheets in the src file')
        log.error('Make sure your workbook worksheets end in the string "Root Synth"')
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
        log.error('Was unable to get any tube data')
        sys.exit(-1)

    # write out the data for each tube into the new worksheet
    foo = write_out_data(options.output, tubes)
    if foo:
        log.info('Sucesfully wrote out data to {}'.format(options.output))
    else:
        log.info('Did not write data out.')
        sys.exit(-1)

    log.info('Done processing all data')
    sys.exit(0)


def root_options():
    parser = argparse.ArgumentParser(prog=__name__)
    parser.add_argument('-s', '--source', dest='src_file', required=True, type=str, action='store',
                        help='Source xlsx file to process')
    parser.add_argument('-o', '--output', dest='output', required=True, type=str, action='store',
                        help='Define the output file (xlsx)')
    parser.add_argument('-v', '--verbose', dest='verbose', default=False, action='store_true',
                        help='Enable verbose output')
    return parser


if __name__ == "__main__":
    p = root_options()
    opts = p.parse_args()

    main(opts)
