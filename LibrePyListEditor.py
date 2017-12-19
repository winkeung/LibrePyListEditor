import sys

# this script can be executed inside Libre Office, using uno or win32com.client (with different initialization code)
try: 
    # #get the doc from the scripting context which is made available to all scripts
    desktop = XSCRIPTCONTEXT.getDesktop()

    import uno
    ctx = XSCRIPTCONTEXT.getComponentContext()
    smgr = ctx.ServiceManager

except:
    try:
        import socket  # only needed on win32-OOo3.0.0
        import uno

        # get the uno component context from the PyUNO runtime
        localContext = uno.getComponentContext()

        # create the UnoUrlResolver
        resolver = localContext.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", localContext)

        # connect to the running office
        ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        smgr = ctx.ServiceManager

        # get the central desktop object
        # desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)
        desktop = smgr.createInstance("com.sun.star.frame.Desktop")

        # access the current writer document
        # model = desktop.getCurrentComponent()
    except:
        # import win32com.client
        import comtypes.client

        # smgr = win32com.client.Dispatch("com.sun.star.ServiceManager")
        smgr = comtypes.client.CreateObject("com.sun.star.ServiceManager")
        desktop = smgr.CreateInstance("com.sun.star.frame.Desktop")

try:
    unicode
except:
    unicode = str

class _Getch:
    """Gets a single character from standard input.  Does not echo to the
screen."""
    def __init__(self):
        try:
            self.impl = _GetchWindows()
        except ImportError:
            self.impl = _GetchUnix()

    def __call__(self): return self.impl()


class _GetchUnix:
    def __init__(self):
        import tty, sys

    def __call__(self):
        import sys, tty, termios
        fd = sys.stdin.fileno()
        old_settings = termios.tcgetattr(fd)
        try:
            tty.setraw(sys.stdin.fileno())
            ch = sys.stdin.read(1)
        finally:
            termios.tcsetattr(fd, termios.TCSADRAIN, old_settings)
        return ch


class _GetchWindows:
    def __init__(self):
        import msvcrt

    def __call__(self):
        import msvcrt
        return msvcrt.getch()


getch = _Getch()

def PythonVersionWordDoc(*args):

    """Prints the Python version into the current document"""
    global desktop
    # get the doc from the scripting context which is made available to all scripts
    # desktop = XSCRIPTCONTEXT.getDesktop()

    model = desktop.getCurrentComponent()
    # check whether there's already an opened document. Otherwise, create a new one
    if not hasattr(model, "Text"):
        model = desktop.loadComponentFromURL(
            "private:factory/swriter", "_blank", 0, ())
    # get the XText interface
    text = model.Text
    # create an XTextRange at the end of the document
    tRange = text.End
    # and set the string
    tRange.String = "Hi, the Python version is %s.%s.%s" % sys.version_info[
                                                           :3] + " and the executable path is " + sys.executable
    return None


def PythonVersionSpreadSheet(*args):
    """Prints the Python version into the current document"""
    global desktop
    # get the doc from the scripting context which is made available to all scripts
    # desktop = XSCRIPTCONTEXT.getDesktop()

    model = desktop.getCurrentComponent()
    # check whether there's already an opened document. Otherwise, create a new one
    if not hasattr(model, "Sheets"):
        model = desktop.loadComponentFromURL(
            "private:factory/scalc", "_blank", 0, ())

    #sheet = model.Sheets.getByIndex(0)  # access the active sheet
    sheet = model.CurrentController.ActiveSheet
    # create an XTextRange at the end of the document
    tRange = sheet.getCellRangeByName("C4")
    # and set the string
    tRange.String = "The Python version is %s.%s.%s" % sys.version_info[:3]
    tRange = sheet.getCellRangeByName("C5")
    tRange.String = sys.executable
    return None

def group(*args):
#     import uno
#     ctx = XSCRIPTCONTEXT.getComponentContext()
#     smgr = ctx.ServiceManager

    model = desktop.getCurrentComponent()

    #VB code
    #    document   = ThisComponent.CurrentController.Frame()
    #    dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

    # access the document
    document = model.getCurrentController()
    # access the dispatcher
    #dispatcher = smgr.createInstanceWithContext( "com.sun.star.frame.DispatchHelper", ctx)
    dispatcher = smgr.createInstance( "com.sun.star.frame.DispatchHelper")

    #VB code
    #    args1 = com.sun.star.beans.PropertyValue()
    #    args1[0].Name = "RowOrCol"
    #    args1[0].Value = "R"
    #    dispatcher.executeDispatch(document, ".uno:Group", "", 0, args1)

    #struct = uno.createUnoStruct('com.sun.star.beans.PropertyValue')
    struct = get_struct()

    struct.Name = 'RowOrCol'
    struct.Value = 'R'

    dispatcher.executeDispatch(document, ".uno:Group", "", 0, tuple([struct]))

    return None

def get_struct():
    try:
        smgr._FlagAsMethod("Bridge_GetStruct")
        struct = smgr.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
    except:
        struct = uno.createUnoStruct('com.sun.star.beans.PropertyValue')
    
    return struct
    
# def group_COM(*args): # merged to group()
    # """ COM verion """

    # model = desktop.getCurrentComponent()

    # document = model.getCurrentController()
    # #dispatcher = OO_ServiceManager.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
    # dispatcher = smgr.CreateInstance("com.sun.star.frame.DispatchHelper")

    # # smgr._FlagAsMethod("Bridge_GetStruct")
    # # struct = smgr.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
    # struct = get_struct()
    
    # struct.Name = "RowOrCol"
    # struct.Value = "R"

    # dispatcher.executeDispatch(document, ".uno:Group", "", 0, tuple([struct]))
    # return None

def set_selection_visible(isVisible):
    model = desktop.getCurrentComponent()
    document = model.getCurrentController()
    dispatcher = smgr.createInstance( "com.sun.star.frame.DispatchHelper")

    struct = get_struct()

    struct.Name = 'RowOrCol'
    struct.Value = 'R'

    if isVisible:
        cmd_str = ".uno:ShowRow"
    else:
        cmd_str = ".uno:HideRow"

    dispatcher.executeDispatch(document, cmd_str, "", 0, tuple([struct]))

def select(scol, srow, lcol, lrow):
    #'dim oSheet, oRange, oCell, oController
    model = desktop.getCurrentComponent()
    oController = model.getCurrentController()
    #oSheet = model.sheets(1)
    #oSheet = model.Sheets.getByIndex(0)  # access the active sheet
    oSheet = model.CurrentController.ActiveSheet
    #oRange = oSheet.getCellRangeByname("B2:D3")
    oRange = oSheet.getCellRangeByPosition(scol, srow, lcol, lrow)
    oController.select(oRange)

def createButton(*args):
    model = desktop.getCurrentComponent()
    sheet = model.Sheets.getByIndex(0)
    #sheet = model.CurrentController.ActiveSheet

    LShape  = model.createInstance("com.sun.star.drawing.ControlShape")

    aPoint = uno.createUnoStruct('com.sun.star.awt.Point')
    aSize = uno.createUnoStruct('com.sun.star.awt.Size')
    aPoint.X = 500
    aPoint.Y = 1000
    aSize.Width = 5000
    aSize.Height = 1000
    LShape.setPosition(aPoint)
    LShape.setSize(aSize)

    oButtonModel = smgr.createInstanceWithContext("com.sun.star.form.component.CommandButton", ctx)
    oButtonModel.Name = "Click"
    oButtonModel.Label = "Python Version"

    LShape.setControl(oButtonModel)

    oDrawPage = sheet.DrawPage
    oDrawPage.add(LShape)

    aEvent = uno.createUnoStruct("com.sun.star.script.ScriptEventDescriptor")
    aEvent.AddListenerParam = ""
    aEvent.EventMethod = "actionPerformed"
    aEvent.ListenerType = "XActionListener"
    aEvent.ScriptCode = "myscript.py$PythonVersionSpreadSheet (user, Python)"
    #aEvent.ScriptCode = "application:Standard.Module1.Main1 (application, Basic)"
    aEvent.ScriptType = "Script"

    oForm = oDrawPage.getForms().getByIndex(0)
    oForm.getCount()
    oForm.registerScriptEvent(0, aEvent)
    return None

def readRange(*arg):
    #get document and so on     
    # cont = uno.getComponentContext()
    # smgr = cont.ServiceManager
    # desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    doc = desktop.getCurrentComponent()
    sheet = doc.CurrentController.getActiveSheet()
   
    #Get user selection
    oSelection = doc.getCurrentSelection()
    oArea = oSelection.getRangeAddress()
    frow = oArea.StartRow
    lrow = oArea.EndRow
    fcol = oArea.StartColumn
    lcol = oArea.EndColumn

    # print(frow, lrow, fcol, lcol)

    if lcol == 1023 or lrow == 1048575 :
        c = sheet.createCursor()
        c.goToEndOfUsedArea(False)

    if lcol == 1023:
        lcol = c.RangeAddress.EndColumn
    if lrow == 1048575:
        lrow = c.RangeAddress.EndRow

    # print(frow, lrow, fcol, lcol)

    #get real range to extract data
    oRange = sheet.getCellRangeByPosition(fcol, frow, lcol, lrow)

    # print(oRange)

    #Extract cell contents as DataArray
    global data_tup
    data_tup = oRange.getDataArray()

    # print type(data_tup)
    # print(data_tup)
    
    return data_tup
    
def writeRange(*arg):
    #get document and so on     
    # cont = uno.getComponentContext()
    # smgr = cont.ServiceManager
    # desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    doc = desktop.getCurrentComponent()
    sheet = doc.CurrentController.getActiveSheet()
   
    #Get user selection
    # oSelection = doc.getCurrentSelection()
    # oArea = oSelection.getRangeAddress()
    frow = 0
    lrow = 4
    fcol = 0
    lcol = 3

    # print(frow, lrow, fcol, lcol)

    #get real range to extract data
    oRange = sheet.getCellRangeByPosition(fcol, frow, lcol, lrow)
    # oRange = sheet.getCellRangeByName("A1:D3")
    
    # #print(oRange)

    # #Extract cell contents as DataArray
    # a = [["1",1.0,3,repr(4)],[5,"6",7,8],[9,10,11,12]]
    a = [
    ['a =', '{', ' ',   ' '],
    [' ',   ' ', 'k3:', 1.0],
    [' ',   ' ', 'k2:', '2'],
    [' ',   ' ', 'k1:', '1'],
    [' ',   '}', ' ',   ' ']
    ]
    # a = ((1,2,3,4),(5,6,7,8),(9,10,11,12))
    # try:
    oRange.setDataArray(a)
    # except:
        # pass
    # and set the string
    # oRange.String = "writeRange"
    
    # print(data_tup)

def rows_to_sheet(rows):
    rows = rectanglize(rows)
    doc = desktop.getCurrentComponent()
    sheet = doc.CurrentController.getActiveSheet()
    oRange = sheet.getCellRangeByPosition(0, 0, len(rows[0])-1, len(rows)-1)
    # tt = []
    # for r in rows:
    #     tt.append(tuple(r))
    # tt = tuple(tt)
    oRange.setDataArray(rows)

def sheet_to_rows():
    doc = desktop.getCurrentComponent()
    sheet = doc.CurrentController.getActiveSheet()

    c = sheet.createCursor()
    c.gotoEndOfUsedArea(False)
    lcol = c.RangeAddress.EndColumn
    lrow = c.RangeAddress.EndRow

    # get real range to extract data
    oRange = sheet.getCellRangeByPosition(0, 0, lcol, lrow)

    # Extract cell contents as DataArray
    global data_tup
    data_tup = oRange.getDataArray()

    return data_tup

def clear_sheet():
    doc = desktop.getCurrentComponent()
    sheet = doc.CurrentController.getActiveSheet()

    c = sheet.createCursor()
    c.gotoEndOfUsedArea(False)
    lcol = c.RangeAddress.EndColumn
    lrow = c.RangeAddress.EndRow

    oRange = sheet.getCellRangeByPosition(0, 0, lcol, lrow)
    oRange.clearContents(0x1ff)

def write_range(rows, fcol, frow):
    doc = desktop.getCurrentComponent()
    sheet = doc.CurrentController.getActiveSheet()
    oRange = sheet.getCellRangeByPosition(fcol, frow, len(rows[0]) + fcol - 1, len(rows) + frow - 1)
    tt = []
    for r in rows:
        tt.append(tuple(r))
    tt = tuple(tt)
    oRange.setDataArray(tt)

# rows assumed to be rectangle shape (i.e. same len for all rows)
def write_to_cell(value, rows, c, r):
    width = len(rows[0])
    while len(rows) < r+1:
        rows.append([''] * width)
    height = len(rows)
    if width < c+1:
        for i in range(height):
            rows[i] = rows[i] + [''] * (c + 1 - width)

    rows[r][c] = value
    return rows

# cur_col = 0 # always point to next available empty cell
# cur_row = 0    # 
# rows = []

def write_obj_old(obj, rows, cur_col, cur_row):
    # global cur_col, cur_row, rows
    if type(obj) == list:
        write_to_cell('[', rows, cur_col, cur_row)
        cur_row = cur_row + 1
        cur_col = cur_col + 1 # indent 
        for i in obj:
            cur_col, cur_row = write_obj(i, rows, cur_col, cur_row)
        cur_col = cur_col - 1 # un-indent    
        write_to_cell(']', rows, cur_col, cur_row)
        cur_row = cur_row + 1
    elif type(obj) == dict:
        write_to_cell('{', rows, cur_col, cur_row)
        cur_row = cur_row + 1
        cur_col = cur_col + 1 # indent 
        for i in obj:
            write_to_cell(repr(i) + ":", rows, cur_col, cur_row) # write key
            cur_col = cur_col + 1
            cur_col, cur_row = write_obj(obj[i], rows, cur_col, cur_row)
            cur_col = cur_col - 1
        cur_col = cur_col - 1 # un-indent    
        write_to_cell('}', rows, cur_col, cur_row)
        cur_row = cur_row + 1
    elif type(obj) == tuple:
        write_to_cell('(', rows, cur_col, cur_row)
        cur_row = cur_row + 1
        cur_col = cur_col + 1 # indent 
        for i in obj:
            cur_col, cur_row = write_obj(i, rows, cur_col, cur_row)
        cur_col = cur_col - 1 # un-indent
        write_to_cell(')', rows, cur_col, cur_row)
        cur_row = cur_row + 1
    elif type(obj) == float:
        write_to_cell(obj, rows, cur_col, cur_row)
        cur_row = cur_row + 1
    else:
        write_to_cell(repr(obj), rows, cur_col, cur_row)
        cur_row = cur_row + 1

    return cur_col, cur_row    

def write_obj(obj, rows, cur_col, cur_row):
    # global cur_col, cur_row, rows
    
    typ = type(obj)
    if typ == list or typ == dict or typ == tuple:
        if typ == list:
            write_to_cell('[', rows, cur_col, cur_row)
        elif typ == dict:
            write_to_cell('{', rows, cur_col, cur_row)
        else:
            write_to_cell('(', rows, cur_col, cur_row)

        cur_row = cur_row + 1
        cur_col = cur_col + 1 # indent
        for i in obj:
            if typ == dict:    
                write_to_cell(repr(i) + ":", rows, cur_col, cur_row) # write key
                cur_col = cur_col + 1
                cur_col, cur_row = write_obj(obj[i], rows, cur_col, cur_row)    
                cur_col = cur_col - 1
            else:
                cur_col, cur_row = write_obj(i, rows, cur_col, cur_row)

        cur_col = cur_col - 1 # un-indent    

        if typ == list:
            write_to_cell(']', rows, cur_col, cur_row)
        elif typ == dict:
            write_to_cell('}', rows, cur_col, cur_row)    
        else:
            write_to_cell(')', rows, cur_col, cur_row)
            
    elif typ == float:
        write_to_cell(obj, rows, cur_col, cur_row)
    else:
        write_to_cell(repr(obj), rows, cur_col, cur_row)

    cur_row = cur_row + 1
    
    return cur_col, cur_row    

def var_to_rows(var, var_name):
    # exec("var=" + var_name)
    rows=[[var_name + "="]]
    write_obj(var, rows, 1, 0)
    rows = list(rows)
    for r in range(len(rows)):
        rows[r] = tuple(rows[r])
    rows = tuple(rows)
    rows = fix_indent(rows)
    return rows

def var_to_sheet(var):
    rows_to_sheet(var_to_rows(var, "x"))

def sheet_to_var():
    rows = sheet_to_rows()
    rows = list(rows)
    rows[0] = (("x="),) + rows[0][1:]
    exec(rows_to_str(rows))
    # print x
    return x

def print_cells(rows):
    max_row_len = 0
    for r in rows:
        max_row_len = max(max_row_len, len(r))

    max_widths = [0] * max_row_len # store max widht for each column
    
    for r in rows: # find max width for each column
        for c in range(len(r)):
            i = r[c]
            if type(i) == str : 
                l = len(repr(i)) - 2 # don't include the quotation mark
            elif type(i) == unicode:
                l = len(repr(i)) - 3 # don't include the unicode mark and quotation mark
            else:
                l = len(repr(i))
            max_widths[c] = max(l, max_widths[c])

    if type(rows) == tuple:
        print ("( ")
    else:
        print ("| ")
    for r in rows:
        if type(r) == tuple:
            s = "( "
        else:
            s = "| "
        for c in range(len(r)):
            i = r[c]
            if type(i) == str : 
                l = len(repr(i)) - 2 # don't include the quotation mark
                s = s + repr(i)[1:-1]
            elif type(i) == unicode:
                l = len(repr(i)) - 3 # don't include the unicode mark and quotation mark
                s = s + repr(i)[2:-1]
            else:
                l = len(repr(i))
                s = s + repr(i)
            s = s + " " * (max_widths[c] - l) + " | "
        print(s)

def clear_range(fcol, frow, lcol, lrow):
    doc = desktop.getCurrentComponent()
    sheet = doc.CurrentController.getActiveSheet()

    oRange = sheet.getCellRangeByPosition(fcol, frow, lcol, lrow)
    oRange.clearContents(0x1ff)

def get_statement():
    doc = desktop.getCurrentComponent()
    sheet = doc.CurrentController.getActiveSheet()

    # Get user selection
    oSelection = doc.getCurrentSelection()
    oArea = oSelection.getRangeAddress()
    frow = oArea.StartRow
    # lrow = oArea.EndRow
    fcol = oArea.StartColumn
    # lcol = oArea.EndColumn

    c = sheet.createCursor()
    c.gotoEndOfUsedArea(False)
    # print("here")
    lcol = c.RangeAddress.EndColumn
    lrow = c.RangeAddress.EndRow

    # read 1 row upward at a time until variable assignment is found
    r = frow
    rows = []
    while True:
        # print(r, lcol)
        oRange = sheet.getCellRangeByPosition(0, r, lcol, r)
        # print("here")
        # Extract cell contents as DataArray
        #global data_tup
        rows = list(oRange.getDataArray()) + rows

        i = rows[0][0]
        # print (i)
        if type(i) == str or type(i) == unicode:
            # print('str')
            if  0 < len(i) :
                # print('0<')
                if i[-1] == '=':
                    # print("found")
                    break

        r = r - 1 # let exception happen if not found until top of sheet is rearched

    var_assign_row = r

    # check if it is a list/dict/tuple
    bracket_lvl = 0
    for c in range(1, len(rows[0])):
        i = rows[0][c]
        if type(i) == str or type(i) == unicode:
            i = i.strip()
            if i == '':
                continue
            elif i == '[' or i == '(' or i == '{':
                bracket_lvl = 1
                break
            else:
                raise Exception('not a list/dict/tuple')
        else:
            raise Exception('not a list/dict/tuple')

    if bracket_lvl == 0:
        raise Exception('not a list/dict/tuple')

    # print bracket_lvl, c, lcol, len(rows)
    # print bracket_lvl, c

    # read forward until the outer most matching closing bracket is found
    rows_idx = 0
    found = False
    while not found:
        # print c, rows_idx
        c = c + 1
        if c == lcol+1:
            c = 0
            rows_idx = rows_idx + 1
        if rows_idx == len(rows):
            break
        i = rows[rows_idx][c]
        if type(i) == str or type(i) == unicode:
            i = i.strip()
            if i == '[' or i == '(' or i == '{':
                bracket_lvl = bracket_lvl + 1
                # print("[ ")
            elif  i == ']' or i == ')' or i == '}':
                bracket_lvl = bracket_lvl - 1
                # print("]")
                if bracket_lvl == 0:
                    return tuple(rows), var_assign_row, fcol, frow

    # print bracket_lvl,c, rows_idx
    # return rows, var_assign_row, fcol, frow

    # read 1 row downward at a time until the outer most matching closing bracket is found
    for r in range(var_assign_row + len(rows), lrow + 1):
        oRange = sheet.getCellRangeByPosition(0, r, lcol, r)

        rows =  rows + list(oRange.getDataArray())

        for i in rows[-1]:
            if type(i) == str or type(i) == unicode:
                i = i.strip()
                if i == '[' or i == '(' or i == '{':
                    bracket_lvl = bracket_lvl + 1
                elif i == ']' or i == ')' or i == '}':
                    bracket_lvl = bracket_lvl - 1
                    if bracket_lvl == 0:
                        found = True
                        break
        if found:
            break

    if not found:
        raise Exception('closing bracket not found')

    return tuple(rows), var_assign_row, fcol, frow

phy = ((1,2,3,'\\'),(),(4,5,6,),(7,8,9,'\\'),(10,11,12),(13,14,15,'\\'))
log = ((1,2,3,4,5,6,7,8,9,10),)
log1 = ((1,2,3,4,5,6,7,8,9,10,'\\'),(11,12,13))

def logical_to_phy_line(rows, max_col):
    """Split locical line(row) into physical lines(rows) by spliting long line
        into shorter lines and add line continutation characters

        Keyword arguments:
        rows -- tuple of tuple(line)
        max_col -- maximum no. of columns allowed in one line

        Return value:
        converted tuple of tuple
    """
    out_rows = ()
    line = ()
    for r in range(len(rows)):
        # search line continuation char from line end backward
        found = False
        for c in range(len(rows[r]) - 1, -1, -1):
            i = rows[r][c]
            if type(i) == str or type(i) == unicode:
                if 0 < len(i):
                    if i[-1] == '\\':
                        found = True
                        break
                    else:
                        break  # string other then '\\' encountered
                else:
                    pass  # blank cell
            else:
                break  # non string data encountered
        else: # loop ended NOT because of break -> it is a blank line
            c = -1

        if found:
            line = line + rows[r][:c] # remove '\\'
        else:
            line = line + rows[r][:c+1] # remove trailing blank cells

        # split if too long
        while max_col < len(line):
            out_rows = out_rows + (line[:max_col-1] + ('\\',),)
            line = line[max_col-1:]

        if found:
            pass
        else:
            out_rows = out_rows + (line,)
            line = ()
    if found:
        if len(line) == max_col:
            out_rows = out_rows + (line[:-1] + ('\\',),)
            out_rows = out_rows + (('\\',),)
        else:
            out_rows = out_rows + (line + ('\\',),)
    return out_rows

def phy_to_logical_line(rows, col, row):
    """Merge pyhsical lines(rows) into logical line(row) by removing line continuation characters and joining the lines.

        Keyword arguments:
        rows -- tuple of tuple(line)
        col, row -- location of a cell in rows

        Return value:
        rows -- all line continuation characters is removed and lines merged.
        col, row -- the new location of the cell after the conversion.
    """
    out_rows = ()
    line = ()
    for r in range(len(rows)):
        # search line continuation char from line end backward
        found = False
        for c in range(len(rows[r]) - 1, -1, -1):
            i = rows[r][c]
            if type(i) == str or type(i) == unicode:
                if  0 < len(i) :
                    if i[-1] == '\\':
                        found = True
                        break
                    else:
                        break # string other then '\\' encountered
                else:
                    pass # blank cell
            else:
                break # non string data encountered

        if r == row:
            out_row = len(out_rows)
            out_col = len(line) + col

        if found:
            line = line + rows[r][:c]
        else:
            line = line + rows[r]
            out_rows = out_rows + (line,)
            line = ()
    if found:
        out_rows = out_rows + (line + ('\\',),)

    return out_rows, out_col, out_row

def get_parent_list_rows():
    """Get the parent list of the element at selection's top-left cell's left edge slice at.

    e.g.  | ( |[a | b]| c | d | ) |  the square represent selection, it slice between cell a and b.

    return value: the rows that cover the list (include brackets)

    """
    doc = desktop.getCurrentComponent()
    sheet = doc.CurrentController.getActiveSheet()

    # Get user selection
    oSelection = doc.getCurrentSelection()
    oArea = oSelection.getRangeAddress()
    frow = oArea.StartRow
    # lrow = oArea.EndRow
    fcol = oArea.StartColumn
    # lcol = oArea.EndColumn

    cursor = sheet.createCursor()
    cursor.gotoEndOfUsedArea(False)
    # print("here")
    lcol = cursor.RangeAddress.EndColumn
    lrow = cursor.RangeAddress.EndRow

    # print(fcol, frow, lcol, lrow)

    r = frow
    rows = ()

    # read upward to find start of logical line
    r = frow - 1
    rows_start_row = frow
    while 0 < r:
        oRange = sheet.getCellRangeByPosition(0, r, lcol, r)
        line = oRange.getDataArray()

        # search line continuation char
        found = False
        for c in range(len(line[0]) - 1, -1, -1):
            i = line[0][c]
            # print (i)
            if type(i) == str or type(i) == unicode:
                # print('str')
                if  0 < len(i) :
                    # print('0<')
                    if i[-1] == '\\':
                        found = True
                        break
                    else:
                        break # string other then '\\' encountered
                else:
                    pass # blank cell
            else:
                break # non string data encountered

        if found:
            rows = line + rows
            rows_start_row = r
        else:
            break

        r = r - 1

    # search forward until bracket_lvl 0 closing bracket is found
    bracket_lvl = 1
    c = fcol + 1
    # read 1 row downward at a time until the outer most matching closing bracket is found
    # print(frow, lrow)
    for r in range(frow, lrow + 1):
        oRange = sheet.getCellRangeByPosition(0, r, lcol, r)
        rows = rows + oRange.getDataArray()

        while c < lcol + 1:
            # print (c,r)
            i = rows[-1][c]
            if type(i) == str or type(i) == unicode:
                i = i.strip()
                if i == '[' or i == '(' or i == '{':
                    bracket_lvl = bracket_lvl + 1
                    # print(i)
                elif  i == ']' or i == ')' or i == '}':
                    bracket_lvl = bracket_lvl - 1
                    # print(i)
                    if bracket_lvl == 0:
                        return rows, rows_start_row, fcol, frow
            c = c + 1
        c = 0

    raise Exception('closing bracket not found')

# # add newline after cell (c,r)
# def newline(rows, c, r):
#     # rows.insert(r+1, rows[r][c+1:] + [''] * (c+1))
#     # rows[r] = rows[r][:c+1] + [''] * (len(rows[r])-c-1)
#     rows.insert(r+1, rows[r][c+1:])
#     rows[r] = rows[r][:c+1]
#     return rows

# add newline after cell (c,r)
def newline(rows, c, r):
    rows = list(rows)
    # rows.insert(r+1, rows[r][c+1:] + [''] * (c+1))
    # rows[r] = rows[r][:c+1] + [''] * (len(rows[r])-c-1)
    rows.insert(r+1, rows[r][c+1:])
    rows[r] = rows[r][:c+1]
    return tuple(rows)

def rstrip_row(row):
    row = list(row)
    for c in range(len(row) - 1, -1, -1):
        i = row[c]
        if i == "":
            row.pop(c)
        else:
            break
    return tuple(row)

def rm_newline(rows, r):
    # merge row r and r+1
    rows = list(rows)

    # remove trailing blank cells in row r
    rows[r] = rstrip_row(rows[r])

    try:
        # remove indentation of row r+1
        for c in range(len(rows[r+1])):
            i = rows[r+1][c]
            if i != '':
                break
        next_row = ()
        if i != '': # 1st non blank cell encountered, otherwise this row is all blank '()'
            next_row = rows[r+1][c:]

        rows[r] = rows[r] + next_row  # copy 'lstrip' of row r+1 to the end of 'rstrip' of row r

        rows.pop(r+1) # remove row r+1
    except:
        pass
    return tuple(rows)

def backward_search_eldest_sibling(rows, c, r):
    bracket_lvl = 0
    for row in range(r - 1, -1, -1):
        for col in range(len(rows[row]) - 1, -1, -1):
            i = rows[r][c]
            if type(i) == str or type(i) == unicode:
                i = i.strip()

                if type(i) == str or type(i) == unicode:
                    i = i.strip()

                    if i != '':
                        continue

                    if i == '[' or i == '(' or i == '{':
                        if bracket_lvl == 0:
                            pass
                    elif i == ']' or i == ')' or i == '}':
                        bracket_lvl = backet_lvl + 1
                        pass
                    elif i[-1] == ":":
                        pass
                        #fonnd
                    else:
                        pass
                else:
                    pass

def fix_indent(rows):
    # find the indent level of the top line, other lines indent relative to this
    #  i.e. found the 1st non blank cell's column no.
    found = False
    for c in range(len(rows[0])):
        i = rows[0][c]
        if type(i) == str or type(i) == unicode:
            i = i.strip()
            if i != '':
                found = True
                break
        else:
            found = True
            break

    if False:
        raise Exception('1st line is a blank line.')

    top_ln_indent_lvl = c
    c += 1
    if i == "{" or i == "(" or i == "[":
        bracket_lvl = 1
    else:
        bracket_lvl = 0

    rows = list(rows)

    # print "row 1", bracket_lvl

    r = 0
    # for r in range(len(rows)):
    # while r < len(rows):
    while True:
        # track bracket level.
        # for c in range(len(rows[r])):
        while c < len(rows[r]):
            i = rows[r][c]
            if type(i) == str or type(i) == unicode:
                i = i.strip()

                if i == '[' or i == '(' or i == '{':
                    bracket_lvl = bracket_lvl + 1
                    # print " [", bracket_lvl
                elif i == ']' or i == ')' or i == '}':
                    bracket_lvl = bracket_lvl - 1
                    # print " ]", bracket_lvl
            c += 1
        c = 0
        r += 1
        if r == len(rows):
            break

        # find 1st non blank cell and do indentation
        found = False
        # for c in range(len(rows[r])):
        while c < len(rows[r]):
            i = rows[r][c]
            if type(i) == str or type(i) == unicode:
                i = i.strip()
                if i != '':
                    found = True
                    break
            else:
                found = True
                break

            c += 1

        if not found: # it is a blank row
            # print ("blank row", c)
            continue

        # 1st non blank cell found, do indentation
        if i == "}" or i == ")" or i == "]":
            # print " )", "r=", r, "c=", c , "bl=", bracket_lvl, "tlil=", top_ln_indent_lvl, rows[r][c:]
            rows[r] = ('',) * (bracket_lvl - 1 + top_ln_indent_lvl) + rows[r][c:]
            # print rows[r]
            c = (bracket_lvl - 1 + top_ln_indent_lvl)
        else:
            # print "r=", r, bracket_lvl
            rows[r] = ('',) * (bracket_lvl + top_ln_indent_lvl) + rows[r][c:]
            c = (bracket_lvl + top_ln_indent_lvl)
        # print bracket_lvl

    return tuple(rows)

# check end of line
# i.e. check (c,r) is the last cell of a row or beyond it are all blank cells
def check_eol(rows, c, r):
    if c == len(rows[r]) - 1:
        return True
    else:
        for i in rows[r][c+1:]:
            # i = i.strip()
            if i != '':
                return False
    return True

def rectanglize(rows):
    rows = list(rows)
    max_row_len = 0
    for r in rows:
        max_row_len = max(max_row_len, len(r))

    for ri in range(len(rows)):
        rows[ri] = rows[ri] + ('',) * (max_row_len- len(rows[ri]))

    return tuple(rows)

# work on cells from selected opening bracket up to corresponding closing bracket
def tree():
    rows, top_row, fcol, frow = get_statement() # get the selected Python variable declaration statement from Calc

    statement_range_lcol = len(rows[0]) - 1
    statement_range_lrow = len(rows) + top_row - 1

    c = fcol
    r = frow - top_row

    rows = newlines(rows, c, r)

    rows = fix_indent(rows)

    clear_range(0, frow, statement_range_lcol, statement_range_lrow) # clear part of org text area that got modified

    # insert new rows
    no_of_insert_row = len(rows) - (statement_range_lrow - top_row + 1)
    if 0 < no_of_insert_row:
        doc = desktop.getCurrentComponent()
        sheet = doc.CurrentController.getActiveSheet()
        sheet.Rows.insertByIndex(statement_range_lrow + 1, no_of_insert_row)

    modify_rows = rectanglize(rows[r:])
    write_range(modify_rows, 0, frow)

    return rows, top_row, fcol, frow

# add newlines after all sibiling at and after the cell (c,r)
def newlines(rows, c, r):
    # turn tuple to list
    rows = list(rows)
    for ri in range(len(rows)):
        rows[ri] = list(rows[ri])

    # print ("!",c, r)

    if c == 0 and r == 0:
        # print repr(rows)
        return rows

    bracket_lvl = 0

    all_already_indented = True
    if rows[r][c] != '' and not check_eol(rows, c, r):
        newline(rows, c, r)
        all_already_indented = False

    c = c + 1

    while r < len(rows):
        while c < len(rows[r]):
            i = rows[r][c]
            if type(i) == str or type(i) == unicode:
                i = i.strip()

                if i == '[' or i == '(' or i == '{':
                    bracket_lvl = bracket_lvl + 1
                    if bracket_lvl == 0:
                        if not check_eol(rows, c, r):
                            newline(rows, c, r)
                elif i == ']' or i == ')' or i == '}':
                    bracket_lvl = bracket_lvl - 1
                    if bracket_lvl < 0:
                        # fix_indent(rows)
                        return rows
                    if bracket_lvl == 0:
                        if not check_eol(rows, c, r):
                            newline(rows, c, r)
                elif i == '':
                    pass
                elif i[-1] == ":":
                    pass
                else:
                    if bracket_lvl == 0:
                        if not check_eol(rows, c, r):
                            newline(rows, c, r)
            else:
                if bracket_lvl == 0:
                    if not check_eol(rows, c, r):
                        newline(rows, c, r)

            c = c + 1

        c = 0
        r = r + 1

    # fix_indent(rows)
    return rows

# work on cells from selected opening bracket up to corresponding closing bracket
def indent_old():
    rows, var_assign_row, fcol, frow = get_statement()

    # column and row of the top-left corner of selection relative to rows[[]], not the actual spread sheet
    rows_c = fcol
    rows_r = frow - var_assign_row

    out_rows = [list(rows[rows_r])[:fcol+1] + [''] * (len(rows[0])-fcol-1)]

    out_rows_c = fcol
    out_rows_r = 0

    bracket_lvl = 0

    start_newline = False # cursor in out_rows need to move down one row (i.e. out_rows_r++)
    is_key_value = False

    i = rows[rows_r][rows_c]
    if type(i) == str or type(i) == unicode:
        i = i.strip()

        if i == '[' or i == '(' or i == '{':
            start_newline = True
            out_rows_c = out_rows_c + 1 # indent
        elif i == ']' or i == ')' or i == '}':
            # TODO set out_rows_c to the col of matching opening bracket
            return
        elif i == '':
            return
        elif i[-1] == ":":
            # out_rows_c = out_rows_c + 1
            is_key_value = True
        else:
            pass
    else:
        pass
    rows_c = rows_c + 1

    while rows_r < len(rows):
        while rows_c < len(rows[rows_r]):
            i = rows[rows_r][rows_c]
            if type(i) == str or type(i) == unicode:
                i = i.strip()

                if i == '[' or i == '(' or i == '{':
                    if bracket_lvl == 0:
                        start_newline = True
                    bracket_lvl = bracket_lvl + 1
                elif i == ']' or i == ')' or i == '}':
                    if bracket_lvl == 0:
                        start_newline = True
                        out_rows_c = out_rows_c - 1
                    bracket_lvl = bracket_lvl - 1
                elif i == '':
                    rows_c = rows_c + 1
                    continue
                elif i[-1] == ":":
                    if bracket_lvl == 0:
                        start_newline = True
                    key_value = True
                else:
                    if bracket_lvl == 0:
                        start_newline = True
            else:
                if bracket_lvl == 0:
                    start_newline = True

            rows_c = rows_c + 1

            if start_newline: # cursor move downward
                out_rows_r = out_rows_r + 1
                start_newline = False
            else: # cursor move right
                out_rows_c = out_rows_c + 1

            write_to_cell(i, out_rows, out_rows_c, out_rows_r)

        rows_c = 0
        rows_r = rows_r + 1
        start_newline = True

    return out_rows

# work on cells from selected opening bracket up to corresponding closing bracket
def indent_all():
    # find the shallowest not completely indented level and indent it 1 level

    rows, var_assign_row, fcol, frow = get_statement()

    while True:
        # Keep trace of the smallest no. , set 1st to a big no.
        # that will possibly be bigger then the actual smallest no.
        min_lvl_not_indented = 65535
        cur_lvl = 0
        cur_bracket_type = ""
        out_rows = [[]]
        for r in range(len(rows)):
            for c in range(len(rows[r])):
                i = rows[r][c]
                if type(i) == str or type(i) == unicode:
                    i = i.strip()

                    if i == '[' or i == '(' or i == '{':
                        cur_lvl = cur_lvl + 1
                        cur_bracket_type = i
                        pass
                    elif i == ']' or i == ')' or i == '}':
                        cur_lvl = cur_lvl - 1
                        if c != cur_lvl:
                            min_lvl_not_indented = min(min_lvl_not_indented, cur_lvl)
                        pass
                    elif i == '':
                        pass
                    elif i[-1] == ":":
                        pass
                    else:
                        if c != cur_lvl:
                            min_lvl_not_indented = min(min_lvl_not_indented, cur_lvl)
                        pass
                else:
                    pass

                pass
        pass

def line():
    # find the deepest not ccompletely unindented levels and unindent it 1 level

    rows, top_row, fcol, frow = get_statement() # get the selected Python variable declaration statement from Calc

    statement_range_lcol = len(rows[0]) - 1
    statement_range_lrow = len(rows) + top_row - 1

    c = fcol
    r = frow - top_row

    rows = rm_newlines(rows, c, r)
    # TODO: check if return rows do not exceed spread sheet limitation for no. of columns

    clear_range(0, frow, statement_range_lcol, statement_range_lrow) # clear part of org text area that got modified

    # insert new rows
    # doc = desktop.getCurrentComponent()
    # sheet = doc.CurrentController.getActiveSheet()
    # sheet.Rows.insertByIndex(statement_range_lrow + 1, len(rows) - (statement_range_lrow - top_row + 1))

    # write modified rows
    modify_rows = rectanglize(rows[r:])
    write_range(modify_rows, 0, frow)

    # delete blank line left
    no_of_row_delete = statement_range_lrow + 1 - frow - len(modify_rows)
    if 0 < no_of_row_delete:
        doc = desktop.getCurrentComponent()
        sheet = doc.CurrentController.getActiveSheet()
        # print (frow + len(modify_rows), statement_range_lrow + 1 - frow - len(modify_rows))
        sheet.Rows.removeByIndex(frow + len(modify_rows),no_of_row_delete )

    return rows, top_row, fcol, frow

def rm_newlines(rows, c, r):
    # turn tuple to list
    rows = list(rows)
    for ri in range(len(rows)):
        rows[ri] = list(rows[ri])

    bracket_lvl = 0

    # all_not_indent = True
    if check_eol(rows, c, r):
        rm_newline(rows, r)
    # else:
    #     all_not_indent = False

    c = c + 1

    while r < len(rows):
        while c < len(rows[r]):
            i = rows[r][c]
            if type(i) == str or type(i) == unicode:
                i = i.strip()

                if i == '[' or i == '(' or i == '{':
                    bracket_lvl = bracket_lvl + 1
                    if bracket_lvl == 0:
                        if check_eol(rows, c, r):
                            rm_newline(rows, r)
                elif i == ']' or i == ')' or i == '}':
                    bracket_lvl = bracket_lvl - 1
                    if bracket_lvl < 0:
                        # fix_indent(rows)
                        return rows
                    if bracket_lvl == 0:
                        if check_eol(rows, c, r):
                            rm_newline(rows, r)
                elif i == '':
                    pass
                elif i[-1] == ":":
                    pass
                else:
                    if bracket_lvl == 0:
                        if check_eol(rows, c, r):
                            rm_newline(rows, r)
            else:
                if bracket_lvl == 0:
                    if check_eol(rows, c, r):
                        rm_newline(rows, r)

            c = c + 1

        c = 0
        r = r + 1

    # fix_indent(rows)
    return rows

def toggle(rows, col):
    # Cycle between all collapse, expand 1 level and expand all levels

    # 1. when no newline found, expand one lvl only
    # 2. when all elements are on individual line, collapse all
    # 3. when not 1 and 2, expand all.

    add_newline_one_lvl = False
    if len(rows) == 1: # expend one level only
        add_newline_one_lvl = True

    # add newlines to every element encountered, if all found already added, then remove them all
    all_already_newline = True
    c = col
    r = 0
    if not check_eol(rows, c, r):
        rows = newline(rows, c, r)
        all_already_newline = False

    c = col + 1
    bracket_lvl = 0
    while r < (len(rows)):
        # print rows[r]
        while c < len(rows[r]):
            i = rows[r][c]
            if type(i) == str or type(i) == unicode:
                i = i.strip()

                if i == '[' or i == '(' or i == '{':
                    bracket_lvl = bracket_lvl + 1
                    if not add_newline_one_lvl or bracket_lvl == 0:
                        if not check_eol(rows, c, r):
                            rows = newline(rows, c, r)
                            all_already_newline = False
                elif i == ']' or i == ')' or i == '}':
                    bracket_lvl = bracket_lvl - 1
                    if bracket_lvl < 0:
                        break
                    if not add_newline_one_lvl or bracket_lvl == 0:
                        if not check_eol(rows, c, r):
                            rows = newline(rows, c, r)
                            all_already_newline = False
                elif i == '':
                    pass
                elif i[-1] == ":":
                    pass
                else:
                    if not add_newline_one_lvl or bracket_lvl == 0:
                        if not check_eol(rows, c, r):
                            rows = newline(rows, c, r)
                            all_already_newline = False
            else:
                if not add_newline_one_lvl or bracket_lvl == 0:
                    if not check_eol(rows, c, r):
                        rows = newline(rows, c, r)
                        all_already_newline = False

            c = c + 1
        if bracket_lvl < 0:
            break
        c = 0
        r = r + 1

    if all_already_newline:  # remove all newlines
        while 1 < len(rows):
            rows = rm_newline(rows, 0)

    # fix_indent(rows)
    return rows

MAX_COL = 20

def toggle_tree():
    rows, sr, c, r = get_parent_list_rows()
    # print_cells(rows)
    rows_len = len(rows)
    rows, c, r = phy_to_logical_line(rows, c, r - sr)
    rows = toggle(rows, c)

    # print_cells(rows)
    rows = fix_indent(rows)
    # print_cells(rows)

    rows = logical_to_phy_line(rows, MAX_COL)
    # print_cells(rows)

    doc = desktop.getCurrentComponent()
    sheet = doc.CurrentController.getActiveSheet()

    # delete rows
    sheet.Rows.removeByIndex(sr, rows_len)
    # print rows_len

    # insert rows
    sheet.Rows.insertByIndex(sr, len(rows))
    # print sr, len(rows)

    write_range(rectanglize(rows), 0, sr)

    # Hide except the 1st and the last physical lines of a logical line which has >=3 physical lines
    total_phy_lines = 0
    for r in range(len(rows)):
        total_phy_lines += 1
        try:
            if rows[r][-1] == "\\":
                pass
            else:
                set_rows_visible(r - total_phy_lines + 2, total_phy_lines - 2, False)
                total_phy_lines = 0
        except: # blank row
            set_rows_visible(r - total_phy_lines + 2, total_phy_lines - 2, False)
            total_phy_lines = 0

    # return rows

def set_rows_visible(start_row, no_of_row, isVisible):
    doc = desktop.getCurrentComponent()
    sheet = doc.CurrentController.getActiveSheet()

    # Col = sheet.Columns[1]
    Rows = sheet.Rows
    for r in range(start_row, start_row + no_of_row):
        try:
            Rows[r].IsVisible = isVisible
            # print("IsVisible")
        except:
            # backup current selection
            oSelection = doc.getCurrentSelection()
            oArea = oSelection.getRangeAddress()
            frow = oArea.StartRow
            lrow = oArea.EndRow
            fcol = oArea.StartColumn
            lcol = oArea.EndColumn

            select(fcol, start_row, fcol, start_row + no_of_row - 1)
            set_selection_visible(isVisible)
            # print("set_selection_visiable")

            # restore previous selection
            select(fcol, frow, lcol, lrow)
            break
    # sheet.Rows.hideByIndex(i,1)
    # Row.IsVisible = False

#    sheet.Rows.insertByIndex(1, 1)

def hide(r):
    doc = desktop.getCurrentComponent()
    sheet = doc.CurrentController.getActiveSheet()

    # Col = sheet.Columns[1]
    Row = sheet.Rows[r]
    # Row = sheet.Rows.get(r)
    # sheet.Rows.hideByIndex(i,1)
    Row.IsVisible = False
    #Row.setVisible(False)
    #sheet.Rows.setIsVisible(1, False)

#    sheet.Rows.insertByIndex(1, 1)

def unhide(r):
    doc = desktop.getCurrentComponent()
    sheet = doc.CurrentController.getActiveSheet()

    # Col = sheet.Columns[1]
    Row = sheet.Rows[r]
    # sheet.Rows.hideByIndex(i,1)
    Row.IsVisible = True

def check_hide(r):
    doc = desktop.getCurrentComponent()
    sheet = doc.CurrentController.getActiveSheet()

    # Col = sheet.Columns[1]
    Row = sheet.Rows[r]
    # sheet.Rows.hideByIndex(i,1)
    return Row.IsVisible


def remove_last_comma(s):
    for i in range(len(s)-1, -1, -1):
        if s[i] == ',':
            return s[:i] + s[i+1:]
        elif s[i] == '\n' or s[i] == '\t' or s[i] == ' ':
            pass
        else:
            return s

# (one bracket occupied one whole cell, when '[]' is not used for index operator)            
def rows_to_str(rows):
    repr_str = ''
    bracket_lvl = 0
    for r in rows:
        for i in r:
            if type(i) == str or type(i) == unicode:
                i = i.strip()
                
                if i == '':
                    pass
                elif i == '[' or i == '(' or i == '{':
                    bracket_lvl = bracket_lvl + 1
                    repr_str = repr_str + i        
                elif i == ']' or i == ')' or i == '}':
                    bracket_lvl = bracket_lvl - 1
                    # repr_str = remove_last_comma(repr_str)
                    
                    repr_str = repr_str + i
                    if 0 < bracket_lvl:
                        repr_str = repr_str + ","
                
                elif i[-1] == ":":
                    repr_str = repr_str + i
                elif i == "\\":
                    repr_str += i
                else:
                    repr_str = repr_str + i
                    if 0 < bracket_lvl:
                        repr_str = repr_str + ","
            else:
                repr_str = repr_str + repr(i)
                if 0 < bracket_lvl:
                    repr_str = repr_str + ","
            repr_str = repr_str + '\t'
        
        repr_str = repr_str.rstrip() + '\n'
        
    return repr_str[:-1]

def hello(*args):
    print("hello, world!")

def excel():
    import win32com.client
    o = win32com.client.Dispatch("Excel.Application")
    o.Visible = 1
    o.Workbooks.Add()
    o.Cells(1,1).Value = "Hello"

a={
    "k1":[1,3,(4,["a", 2, 3.0],5),6,7], 
    "k2":"abc", 
    "k3":3, 
    "k4":2.0,
    1234:1,
    (1,2,3,4,):4,
}    

b=[
[1,2,
 3,4],
[1,2,3,4],
[1,2,3,4],
[1,2,3,4]
]
c =[1,2,3]

if __name__ == "__main__":
    toggle_tree()