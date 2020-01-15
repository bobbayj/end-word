# Written by Reddit user _DTR_
import re
import xml.etree.ElementTree as ET
import zipfile

# The following prefixes are prepended to xml tags within xlsx files.
# Make our lives easier and give them their own variables
PREFIX = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
REL = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'

class Workbook:
    '''
    Class that contains the overarching workbook. Very simple at the moment, containing only a
    dictionary of sheets and the currently active one
    '''
    def __init__(self):
        self.sheets = {}
        self.active = None

    def add_sheet(self, sheet, name, active=False):
        self.sheets[name] = sheet
        if active:
            self.active = sheet

class Worksheet:
    '''
    The worksheet class that holds the cell values
    '''
    def __init__(self):
        self.cells = {}
        self.def_cell = SharedString()

        # Keep things 1-based
        self.dim = { 'rw_first' : 1, 'rw_last' : 1, 'col_first' : 1, 'col_last' : 1 }

    def add_cell(self, location, string):
        '''
        Adds a cell at the given A1 reference
        '''
        # self.cells[location] = string
        rw, col = CellHelpers.rwcol_from_ref(location)
        if not rw in self.cells:
            self.cells[rw] = {}

        self.cells[rw][col] = string

        self.dim['rw_first'] = min(rw, self.dim['rw_first'])
        self.dim['rw_last'] = max(rw, self.dim['rw_last'])
        self.dim['col_first'] = min(col, self.dim['col_first'])
        self.dim['col_last'] = max(col, self.dim['col_last'])

    def cell(self, rw, col):
        '''
        Retrieve a cell at the given (1-based) row and column
        '''
        return CellHolder(rw, col, self._cell(rw, col))

    def get_range(self, rng, row_major=True):
        '''
        Retrieve a range of cells determined by the A1 reference string rng
        By default, cells are gathered left-to-right, top-to-bottom (row-major order)
        To iterate by columns first (top-down, left-to-right), set row_major to false
        Handled strings:
            - Single cell ('A4', 'XFD10000') - returns a list of length one, containing the given cell
            - General range ('A1:B2') - returns a two-dimensional list of all the cells in the given range
                - If the range is within a single row or single column, return a one dimensional list,
                otherwise a 2 dimensional list (apaplies to the below as well)
            - Entire columns ('A:A', 'C:Z') - returns an array of all cells in the specified rows, starting
            at row 1 to the row of the last non-blank cell in the worksheet
            - Entire rows ('1:1', '3:5') - same as columns, swapped
        '''
        rng = rng.upper()
        sep_index = rng.find(':')
        if sep_index == -1:
            rw, col = _rw_col(rng)
            return [CellHolder(rw, col, self._cell(rw, col))]

        rw_first, rw_last, col_first, col_last = 4 * [1]

        # Need to check for entire rows/cols. These checks aren't foolproof, they aren't really meant to be
        # Assuming valid input, they do the right thing
        if not re.match(r'\w+\d', rng[:sep_index]):
            if (re.match(r'\w+\d', rng[sep_index + 1:])):
                # invalid, can't have something like A:A1 or A:B5
                return []
            rw_first = 1
            rw_last = self.dim['rw_last']
            col_first = CellHelpers.col_to_num(rng[:sep_index])
            col_last = CellHelpers.col_to_num(rng[sep_index + 1:])
        elif re.match(r'\d+:\d+', rng):
            # Entire column
            rw_first = int(rng[:sep_index])
            rw_last = rw_int(rng[sep_index + 1:])
            col_first = 1
            col_last = self.dim['col_last']
        else:
            rw_first, col_first = CellHelpers.rwcol_from_ref(rng[:sep_index])
            rw_last, col_last = CellHelpers.rwcol_from_ref(rng[sep_index + 1:])

        # Invalid range, return and empty list
        if col_first > col_last or rw_first > rw_last:
            return []

        res = []
        if rw_first == rw_last or col_first == col_last:
            # In the single row/column case, don't return a nested list, just a single one
            for rw in range(rw_first, rw_last + 1):
                for col in range(col_first, col_last + 1):
                    res.append(CellHolder(rw, col, self._cell(rw, col)))
        elif row_major:
            for rw in range(rw_first, rw_last + 1):
                rw_vals = []
                for col in range(col_first, col_last + 1):
                    rw_vals.append(CellHolder(rw, col, self._cell(rw, col)))

                res.append(rw_vals)
        else:
            for col in range(col_first, col_last + 1):
                col_vals = []
                for rw in range(rw_first, rw_last + 1):
                    col_vals.append(CellHolder(rw, col, self._cell(rw, col)))
                res.append(col_vals)

        return res

    def _cell(self, rw, col):
        '''
        Internal method to return the value at a given cell (or a default blank cell if there is no value)
        '''
        if rw in self.cells and col in self.cells[rw]:
            return self.cells[rw][col]
        return self.def_cell

class CellHelpers:
    '''
    Class containing helper methods for converting between row/col and A1 reference styles
    '''
    def col_to_num(self, col):
        '''
        Convert the given A1 column to a 1-based index
        '''
        if len(col) == 1:
            return ord(col) - ord('A') + 1
        if len(col) == 2:
            return 26 + (26 * (ord(col[0]) - ord('A'))) + (ord(col[1]) - ord('A')) + 1
        else:
            return 702 + (676 * (ord(col[0]) - ord('A'))) + (26 * (ord(col[1]) - ord('A'))) + (ord(col[2]) - ord('A')) + 1

    def rwcol_from_ref(self, ref):
        '''
        Convert a single-cell A1 reference to a row and column
        '''
        index = 0
        while ord(ref[index]) < ord('0') or ord(ref[index]) > ord('9'):
            index += 1
        return int(ref[index:]), self.col_to_num(ref[:index])

    def num_to_col(self, num):
        '''
        Convert a column index into an A1 reference
        '''
        num -= 1 # We expect a 1-based col, make it 0-based before doing math
        val = ""
        if num >= 702:
            val = chr(((num - 702) // 676) + ord('A'))
            num %= 676

        if num >= 26:
            val += chr(((num - 26) // 26) + ord('A'))
            num %= 26

        return val + chr(num + ord('A'))

    def a1(self, rw, col):
        '''
        Convert a row and column index to an A1 reference string
        '''
        return self.num_to_col(col) + str(rw)

    def build_range(self, rw_first, rw_last, col_first, col_last):
        '''
        Builds a reference string within the four conrer bounds
        '''
        return self.num_to_col(col_first) + str(rw_first) + ":" + self.num_to_col(col_last) + str(rw_last)


class CellHolder:
    '''
    Simple wrapper for a cell value
    '''
    def __init__(self, rw, col, value):
        self.value = value
        self.rw = rw
        self.col = col

    def __repr__(self):
        return self.value.plain_text()

class SharedString:
    '''
    SharedString contains the contents of a single cell. The string may contain
    multiple runs each with different formatting
    '''
    def __init__(self, text=None, properties=None):
        self.runs = []
        if text != None:
            self.runs.append(Run(plain=text, properties=properties))

    def add_run(self, xmlNode):
        self.runs.append(Run(xmlNode=xmlNode))

    def plain_text(self):
        '''
        Return a plain-text representation of the cell value
        '''
        return ''.join([run.to_string() for run in self.runs])

    def replace(self, find, rep):
        '''
        Replace text within the string. Only supports non-spanning find/replace, i.e. if the find string
        spans changes in formatting, it is not replaced
        '''
        for run in self.runs:
            run.text = run.text.replace(find, rep)

    def add_to_paragraph(self, p):
        '''
        Add this string to the given docx paragraph. Probably shouldn't be a part of this class
        '''
        for run in self.runs:
            docRun = p.add_run(run.to_string())

            # For now, only look at super/subscript, bold, underline, and italic
            if run.has_attr('vertAlign'):
                if run.attrib('vertAlign') == 'subscript':
                    docRun.font.subscript = True
                elif run.attrib('vertAlign') == 'superscript':
                    docRun.font.superscript = True
            if run.has_attr('b'):
                docRun.font.bold = True
            if run.has_attr('u'):
                docRun.font.underline = True
            if run.has_attr('i'):
                docRun.font.italic = True

    def __str__(self):
        return self.plain_text()

    def __repr__(self):
        return self.plain_text()

class Run:
    '''
    A single (un)formatted run of text.

    Input is either an xml node or plain text. If it's an xml node, it is parsed
    and formatting properties are applied. If it's plain text, add it directly
    without any properties
    '''
    def __init__(self, xmlNode=None, plain=None, properties=None):
        self.text = ''
        self.properties = {}

        if plain != None:
            self.text = plain
            if properties != None:
                self.properties = properties.copy()
            return

        # xmlNode  better not be none!
        if xmlNode.tag == PREFIX + 't':
            # Easy case, a single string
            self.text = xmlNode.text
            return
        elif xmlNode.tag != PREFIX + 'r':
            print("Unknown tag: " + xmlNode.tag[len(PREFIX):])
            return

        for child in xmlNode:
            if child.tag == PREFIX + 't':
                self.text = child.text
            elif child.tag != PREFIX + 'rPr':
                print("Unknown flag: ", child.tag[len(PREFIX):])
                continue
            for prop in child:
                item = prop.tag[len(PREFIX):]
                if len(prop.attrib) == 0:
                    # Some properties (bold/italic/underline/etc) don't have attributes, but we should still get the tag
                    self.properties[item] = True
                else:
                    for attr in prop.attrib:
                        self.properties[item] = prop.attrib[attr]
                        break;

    def to_string(self):
        return self.text

    def attributes(self):
        return self.properties

    def attrib(self, attribute):
        return self.properties[attribute] if attribute in self.properties else None

    def has_attr(self, attr):
        return attr in self.properties

def custom_load_workbook(source_file):
    '''
    Reads in the given workbook file and returns a Workbook object containing its sheets and cell values
    '''

    # This assumes an xlsx file that has all the required parts
    container = zipfile.ZipFile(source_file)
    stringFile = ET.parse(container.open('xl/sharedStrings.xml'))


    # Build up our list of shared strings
    strings = []
    for child in stringFile.getroot():
        text = SharedString()
        for run in child:
            text.add_run(run)
        strings.append(text)

    # Build up our list of styles
    cell_styles = []
    if 'xl/styles.xml' in container.namelist():
        style = ET.parse(container.open('xl/styles.xml'))
        fonts_xml = style.getroot().find(PREFIX + 'fonts')
        fonts = []
        for font in fonts_xml:
            props = {}
            for prop in font:
                item = prop.tag[len(PREFIX):]
                if len(prop.attrib) == 0:
                    props[item] = True
                else:
                    props[prop.tag[len(PREFIX):]] = prop.attrib[list(prop.attrib.keys())[0]]
            fonts.append(props.copy())

        cellxfs = style.getroot().find(PREFIX + 'cellXfs')
        index = 0
        for xf in cellxfs:
            cell_styles.append({})
            # Only care about font properties for now
            if 'fontId' in xf.attrib:
                fid = int(xf.attrib['fontId'])
                cell_styles[index]['font'] = fonts[fid]
            else:
                cell_styles[index]['font'] = {}
            index += 1


    wb = Workbook()

    # Get the friendly names of the sheets
    sheetNames = {}
    workbook = ET.parse(container.open('xl/workbook.xml'))
    elements = workbook.getroot().findall(PREFIX + 'sheets/' + PREFIX + 'sheet')
    for sheet in elements:
        sheetNames[sheet.attrib['sheetId']] = { 'name' : sheet.attrib['name'] }

    # now go through our worksheets and find matching string entries
    xmlSheets = [file for file in container.namelist() if file[0:16] == 'xl/worksheets/sh']
    for xmlSheet in xmlSheets:
        ws = Worksheet()
        name = sheetNames[xmlSheet[xmlSheet.find('/sheet') + 6:xmlSheet.rfind('.')]]['name']
        active = False
        if xmlSheet.find(PREFIX + '/sheetViews' + PREFIX + '/sheetView[@tabSelected="1"]'):
            active = True
        tree = ET.parse(container.open(xmlSheet))
        data = tree.getroot().find(PREFIX + 'sheetData')
        cells = data.findall(PREFIX + 'row/' + PREFIX + 'c')
        for cell in cells:
            cell_value = None
            # Cell attributes:
            # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cell?view=openxml-2.8.1
            if ('t' in cell.attrib and cell.attrib['t'] == 's'):
                # We have a shared string
                ws.add_cell(cell.attrib['r'], strings[int(cell.find(PREFIX + 'v').text)])
            else:
                # otherwise we just have the value.
                properties = None
                if ('s' in cell.attrib):
                    style_id = int(cell.attrib['s'])
                    properties = cell_styles[style_id]['font']
                
                # Format numbers as desired
                try:
                    cell_value = cell.find(PREFIX + 'v').text
                    cell_value = f'{float(cell_value):.2f}'
                except AttributeError as e:
                    # Will catch blank cells, i.e. merged cells
#                     print(f'Cell {cell.attrib["r"]} is blank...skipping\n')
                    pass
                
                if cell_value:
                    ws.add_cell(cell.attrib['r'], SharedString(cell_value, properties=properties))
        wb.add_sheet(ws, name, active)
    
    return wb
