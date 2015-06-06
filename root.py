'''
root.py

library for handling root data from winrhizotron

'''
import openpyxl
import logging
import itertools


# logging config
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s: %(levelname)s: %(message)s [%(filename)s:%(funcName)s]')

class Root:
    def __init__(self, rootName, location, birthSession):
        self.rootName = rootName
        self.location = location
        self.tube = ''
        self.session = ''
        self.birthSession = birthSession
        self.goneSession = ''
        self.anomaly = ''
        self.order = ''
        self.highestOrder = ''
        self.censored = ''
        self.aliveTipsAtBirth = ''
        self.aliveTipsAtGone = ''
        self.avgDiamter = ''
        self.isAlive = ''
        self.identity = (rootName, location, birthSession)

    def get_identity(self):
        return (self.rootName, self.location, self.birthSession)

    def set_order(self, order):
        pass

    def set_highest_order(self, highestOrder):
        pass

class Tube:
    def __init__(self, tubeNumber):
        self.tubeNumber = tubeNumber
        self.roots = []
        self.index = 0
    def __iter__(self):
        return self
    def next(self):
        if self.index == len(self.roots):
            self.index = 0
            raise StopIteration
        root = self.roots[self.index]
        self.index = self.index + 1
        return root
    def add_root(self, root):
        root.tube = self.tubeNumber
        self.roots.append(root)
    def insert_or_update_root(self, root):
        # insert root if the root identity is new
        insert = True
        for existingRoot in self:
            if root.identity == existingRoot.identity:
                logging.debug('Root already exists in tube')
                print root.identity
                print existingRoot.identity
                # should update root with new information if possible
                insert = False
        # add the root to the tube
        if insert:
            logging.debug('Adding root to tube')
            print root.identity
            self.add_root(root)
        return True

def get_test_ws():
    wb = openpyxl.load_workbook(filename = r'unedited.xlsx')
    ws = wb.get_sheet_by_name('T001 Root Synth')
    return ws

''' general helper function for ws row/column access'''

def get_index_by_value(cell_list, value):
    ''' given a list of cells and a value, return the index in the list of
    that cell value.  this returns the index of the first time the value
    is found.

    if the value is not found, return false
    '''
    index = False
    for cell in cell_list:
        if cell.value == value:
            try:
                index = cell_list.index(cell)
            except ValueError, e:
                logging.error('Failed to get index for value after finding a match')
                logging.error('%s' %(e))
                return index
            break
    return index

def get_index_at_null(cell_list):
    ''' get the index value for the first NoneType object present in cell list'''
    index = False
    for cell in cell_list:
        if not cell.value:
            index = cell_list.index(cell)
            break
    return index

''' helper functions for finding the import values from the root data table'''

def build_root_data_table(ws):
    ''' BUILD A LIST OF DATA REPRESENTING EACH ROW IN WORKSHEET UP
    TO THE SYNTHESIS TABLE

    Values are chosen by MSS.
    '''
    # build a dict for header values that are important.
    # session numbers map to dates
    headerIndexData = {
    'RootName': 0,
    'Location#': 0,
    'Session#': 0,
    'GoneSession': 0,
    'BirthSession': 0,
    'TotAvgDiam(mm/10)': 0,
    'TipLivStatus': 0,
    'RootNotes': 0,
    'HighestOrder': 0,
    'NumberOfTips': 0,
    }
    rootData = []
    # get index values for each important column
    tableHeader = ws.rows[0]
    for key in headerIndexData:
        index = get_index_by_value(tableHeader, key)
        if not index:
            logging.error('Failed to obtain header value: %s' %(key))
            return False
        headerIndexData[key] = index
    endOfRootData = get_index_at_null(ws.columns[0])
    if not endOfRootData:
        logging.error('Failed to find end of root data.')
        return False
    # this slice includes the header row in rootData.  This lets us
    # rebuild indices, since the dict iterator is not sorted
    for row in ws.rows[:endOfRootData]:
        rootRow = []
        for key in headerIndexData:
                value = row[headerIndexData[key]].value
                rootRow.append(value)
        rootData.append(rootRow)
    return rootData


def process_worksheet(ws):
    '''Process a worksheet, and build a tube object containing all the root data.
    This data can then be dumped out when all tubes have been processed.'''
    rootIndexData = {
    'RootName': 0,
    'Location#': 0,
    'Session#': 0,
    'BirthSession': 0,
    'TotAvgDiam(mm/10)': 0,
    'TipLivStatus': 0,
    'RootNotes': 0,
    'NumberOfTips': 0,
    }
    tubeNumber = get_tube_number(ws)
    tube = Tube(tubeNumber)
    root_data = build_root_data_table(ws)
    if not root_data:
        logging.error('Failed to get root data from sheet: %s' % (ws.title))
        return False
    root_tip_data = build_synthesis_data_table(ws)
    if not root_data:
        logging.error('Failed to get root tip data from sheet: %s' % (ws.title))
        return False
    # process root data to get the roots. root_data[0] should have header root
    headerRow = root_data[0]
    for key in rootIndexData:
        try:
            index = headerRow.index(key)
        except ValueError, e:
            logging.error('Failed to obtain header value: %s' %(key))
            logging.error('%s' % (e))
            return False
        rootIndexData[key] = index
    # this loop will not handle exceptional roots, it just builts a list of root
    # object for each row of root data
    roots = []
    for row in root_data[1:]:
        tempRoot = root_from_row(row, rootIndexData)
        roots.append(tempRoot)
    logging.debug('Identified %s roots in sheet: %s' % (str(len(roots)),ws.title))
    # EXCEPTION ROOT HANDLING WOULD WOULD GO HERE
    # add each root to the tube
    for root in roots:
        tube.insert_or_update_root(root)
    logging.debug('Identified %s roots in sheet: %s' % (str(len(roots)),ws.title))
    logging.debug('Found %s roots in tube' % (str(len(tube.roots))))
    return tube

def root_from_row(row, indexDict):
    ''' build a root object from row, using indexDict for mappings'''
    # identification information
    rootName = row[indexDict['RootName']]
    location = row[indexDict['Location#']]
    birthSession = row[indexDict['BirthSession']]
    root = Root(rootName, location, birthSession)
    # check to see if anomalous root
    if row[indexDict['NumberOfTips']] == 1:
        root.anomaly = False
    else:
        root.anomaly = True
    # root metadata
    root.session = row[indexDict['Session#']]
    root.isAlive = row[indexDict['TipLivStatus']]
    if not root.anomaly:
        if root.isAlive.startswith('G'):
            root.goneSession = root.session
    # order & diameter data
    root.order = row[indexDict['RootNotes']]
    root.avgDiameter = row[indexDict['TotAvgDiam(mm/10)']]
    return root

def get_tube_number(ws):
    ''' get tube number from a given worksheet.  this assumes that '''
    row1 = ws.rows[0]
    tubeIndex = get_index_by_value(row1, 'Tube#')
    logging.debug('TubeIndex: %s' %(str(tubeIndex)))
    tubeNumber = ws.columns[tubeIndex][1].value
    if tubeNumber:
        return tubeNumber
    else:
        return False

''' helper function for working with openpyxl'''
def get_root_sheets(wb):
    '''given a workbook, return all sheet names that end in "Root Synth""'''
    root_sheets = []
    sheets = wb.get_sheets_names()
    for sheet in sheets:
        if sheet.endswith('Root Synth'):
            root_sheets.append(sheet)
    return root_sheets

''' helper functions for finding & using the synthTable from a worksheet'''

def build_synthesis_data_table(ws):
    ''' BUILD A LIST OF DATA REPRESENTING THE TIP TABLE

    header columns are choosen by MSS.  This data is intended to be
    associated with a Tube class.
    '''
    headerIndexData = {'Location#': 0, 'BirthSession' : 0, 'AliveTipsAtBirth': 0}
    tipData = []

    columnA = ws.columns[0]
    synthHeaderRow = find_tip_synthesis_row(columnA)

    # get index location of the important synthesis data tables
    synthHeader = ws.rows[synthHeaderRow]
    for key in headerIndexData:
        index = get_index_by_value(synthHeader, key)
        if not index:
            logging.error('Failed to get synthHeader: %s' % (key))
            return False
        headerIndexData[key] = index

    # build a list of the tip data values
    for row in ws.rows[synthHeaderRow:]:
        location = row[headerIndexData['Location#']].value
        birthSession = row[headerIndexData['BirthSession']].value
        aliveTips = row[headerIndexData['AliveTipsAtBirth']].value
        tipData.append((location, birthSession, aliveTips))

    return tipData

def find_tip_synthesis_row(column):
    synthTableRow = None
    for cell in column:
        # the second clause in the OR statement is done to account for
        # something MSS did while manipulating part of the test data
        if cell.value == ('RootName' or 'Tube#'):
            synthTableRow = column.index(cell)
            logging.debug('Found Synthtable at %s')
    return synthTableRow

def find_tip_synthesis_width(row):
    # this function may not be neccesary
    synthTableWidth = None
    for cell in row:
        if cell.value:
           synthTableWidth = row.index(cell)
        if not cell.value:
            break
    if not synthTableWidth:
        return False
    else:
        return synthTableWidth

def stats(wb, sheetlist, value, verbose=False):
    '''generate counts of different values in a given column, based on column name
    this is done across a workbook'''
    dicty = {}
    for sheet in sheetlist:
            ws = wb.get_sheet_by_name(sheet)
            row1 = ws.rows[0]
            index = get_index_by_value(row1,value)
            for cell in ws.columns[index]:
                    if cell.value in dicty:
                        dicty[cell.value] = dicty[cell.value] + 1
                    else:
                        dicty[cell.value] = 1
                        if verbose:
                            logging.info('Found new cell.value in sheet: %s' % (sheet))
    return dicty