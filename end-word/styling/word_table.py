from docx.oxml import OxmlElement  # For defining and targeting xml elements to change
from docx.oxml.ns import qn  # For defining and targeting xml elements to change

def style_tbl(table, xls_formats):
    tbl = table._tbl # get xml element of the table
    merged_cell_flag = False  # initiate merged cell flag to mark if cell is part of marged group
    merged_cell_count = 0  # initiate counter for merged cells
    
    for cell in tbl.iter_tcs():
        coord = (cell.bottom-1, cell._grid_col)
        tcPr = cell.tcPr  # Get table cell properties
        merged_cell_flag, merged_cell_count = check_merge(tcPr, merged_cell_count, merged_cell_flag)
        
        # FOR TESTING: If it is part of a merged cell, go to next cell in loop
#         if merged_cell_flag:
#             print(f'Skipped cell {coord}')
#             continue
        
        # Run style changes
        _borders(cell, tcPr, xls_formats[coord])
        _fill_align(cell, tcPr, xls_formats[coord])
        _fonts(cell,  tcPr, xls_formats[coord])


def _borders(cell, tcPr, xls_format):
    tcBorders = OxmlElement('w:tcBorders')
    for position in ['top', 'bottom', 'left', 'right']:
            
        # Map xls border format to xml format
        if xls_format['border'][position] == 'thin':
            val = 'single'
            sz = '2'
        elif xls_format['border'][position] == 'medium':
            val = 'single'
            sz = '8'
        elif xls_format['border'][position] == 'thick':
            val = 'single'
            sz = '16'
        else:
            val = 'nil'
            sz = '2'
        
        # Set border formats on obj
        # More options at http://officeopenxml.com/WPtableBorders.php
        side = OxmlElement(f'w:{position}')
        side.set(qn('w:val'), val)
        side.set(qn('w:sz'), sz)  # sz 2 = 1/4 pt is the minimum
        side.set(qn('w:space'), '0')
        side.set(qn('w:shadow'), 'false')
        if xls_format['border'][f'{position}Color'] is not None:  # Catch Nonetype colors
            side.set(qn('w:color'), xls_format['border'][f'{position}Color'])
            
        tcBorders.append(side)
    tcPr.append(tcBorders)
    

def _fill_align(cell, tcPr, xls_format):
    # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.shading?view=openxml-2.8.1
    # Set cell fill 
    fillshade = OxmlElement('w:shd')
    fillshade.set(qn('w:fill'), xls_format['fillColor'])
    tcPr.append(fillshade)
    
    # Set alignment
    vAlign = OxmlElement('w:vAlign')
    try:
        vAlign.set(qn('w:val'), xls_format['vertical'])
        tcPr.append(vAlign)
    except TypeError:
#             print(f'No vertical alignment for a table @ cell {coord} - skipping...')
        pass
    
    
def _fonts(cell, tcPr, xls_format):
    # https://python-docx.readthedocs.io/en/latest/dev/analysis/features/text/font-color.html
    try:
        run = cell.p_lst[0].r_lst[0]
    except IndexError:
        # print(f'IndexError: No run in cell {coord} - skipping')
        return
    rPr = run._add_rPr()
    # Set font color
    if xls_format['fontColor']:
        fontColor = OxmlElement('w:color')
        fontColor.set(qn('w:val'), xls_format['fontColor'])
        rPr.append(fontColor)
    # Set bold
    if xls_format['bold']:
        fontBold = OxmlElement('w:b')
        rPr.append(fontBold)
    # Set font size
    if xls_format['size']:
        size_val = xls_format['size'] * 2  # Measurements are in half-points
        size_val = str(int(size_val))
        fontSize = OxmlElement('w:sz')
        fontSize_cs = OxmlElement('w:szCs')
        fontSize.set(qn('w:val'), size_val)
        fontSize_cs.set(qn('w:val'), size_val)
        rPr.append(fontSize)
        rPr.append(fontSize_cs)
    # Set font name
    if xls_format['name']:
        fontName = OxmlElement('w:rFonts')
        fontName.set(qn('w:ascii'),xls_format['name'])
        fontName.set(qn('w:hAnsi'),xls_format['name'])
        fontName.set(qn('w:cs'),xls_format['name'])
        rPr.append(fontName)

def check_merge(tcPr, merged_cell_count, merged_cell_flag):
    # Note: merged_cell_flag is set to false for the first cell in a merged group
    # If the cell is the start of a new merge, set flags to true and reset merge counter
    if tcPr.gridSpan is not None:
        merged_cell_count = tcPr.grid_span
    # Else if the cell is merged, reduce merged_cell_count by 1
    elif merged_cell_count > 0:
        merged_cell_count -= 1
        merged_cell_flag = True
    # Else if the cell is no longer merged, set merge cell flag to false
    elif merged_cell_count == 0:
        merged_cell_flag = False
    else:
        raise ValueError('Something went wrong with checking merged cells...This condition should never be hit')
    return merged_cell_flag, merged_cell_count
