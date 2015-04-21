#
# Copyright (c) 2011 Andreas Stricker <andy@knitter.ch>
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.
#

import xml.parsers.expat
import zipfile
import datetime
import re
import xldate

class DocumentArchive(object):
    """
    Represent the document ZIP archive. Provide accessor for file handles
    to fetch the archive files.
    """
    def __init__(self, filename):
        self.zip_filename = filename
        self.zip_filehandle = zipfile.ZipFile(filename)

    def filehandle(self, name):
        return self.zip_filehandle.open(name, 'r')

    def workbook(self):
        return self.filehandle('xl/workbook.xml')

    def sheet(self, id):
        return self.filehandle('xl/worksheets/' + self.sheet_filename(id))

    def sheet_filename(self, sheet_id):
        return u'sheet%s.xml' % sheet_id

    def shared_strings(self):
        return self.filehandle('xl/sharedStrings.xml')

    def styles(self):
        return self.filehandle('xl/styles.xml')

class SharedStrings(list):
    """ Parse the shared string list and store it as list for lookup. """
    def __init__(self, archive):
        parser = xml.parsers.expat.ParserCreate()
        parser.StartElementHandler = self._start_element
        parser.EndElementHandler = self._end_element
        parser.CharacterDataHandler = self._char_data

        self.is_string = False
        self.data = None
        self.index = 0

        fh = archive.shared_strings()
        parser.ParseFile(fh)
        fh.close()

    def _start_element(self, name, attrs):
        if name == 'si':
            self.is_string = True

    def _end_element(self, name):
        self.is_string = False
        if name == 'si':
            self.append(self.data)
            self.data = None

    def _char_data(self, data):
        if self.is_string is True:
            if self.data is None:
                self.data = data
            else:
                self.data += data

class Styles(object):
    """ Parse style file and provide methods to work with styles and formats. """
    def _dateWrapper(*args):
        if len(args) == 6:
            return datetime.datetime(*args)
        else:
            return datetime.date(*args)
    BUILTIN_FMT = 0
    BUILTIN_TYPE = 1
    # Stolen from perls Spreadsheet::XLSX
    BUILTIN_NUM_FMTS = {
        0x00: ('@', unicode),
        0x01: ('0', int),
        0x02: ('0.00', float),
        0x03: ('#,##0', float),
        0x04: ('#,##0.00', float),
        0x05: ('($#,##0_);($#,##0)', float),
        0x06: ('($#,##0_);[RED]($#,##0)', float),
        0x07: ('($#,##0.00_);($#,##0.00_)', float),
        0x08: ('($#,##0.00_);[RED]($#,##0.00_)', float),
        0x09: ('0%', int),
        0x0A: ('0.00%', float),
        0x0B: ('0.00E+00', float),
        0x0C: ('# ?/?', float),
        0x0D: ('# ??/??', float),
        0x0E: ('m-d-yy', _dateWrapper),
        0x0F: ('d-mmm-yy', _dateWrapper),
        0x10: ('d-mmm', _dateWrapper),
        0x11: ('mmm-yy', _dateWrapper),
        0x12: ('h:mm AM/PM', datetime.time),
        0x13: ('h:mm:ss AM/PM', datetime.time),
        0x14: ('h:mm', datetime.time),
        0x15: ('h:mm:ss', datetime.time),
        0x16: ('m-d-yy h:mm', datetime.datetime),
#0x17-0x24 -- Differs in Natinal
        0x25: ('(#,##0_);(#,##0)', int),
        0x26: ('(#,##0_);[RED](#,##0)', int),
        0x27: ('(#,##0.00);(#,##0.00)', float),
        0x28: ('(#,##0.00);[RED](#,##0.00)', float),
        0x29: ('_(*#,##0_);_(*(#,##0);_(*"-"_);_(@_)', float),
        0x2A: ('_($*#,##0_);_($*(#,##0);_(*"-"_);_(@_)', float),
        0x2B: ('_(*#,##0.00_);_(*(#,##0.00);_(*"-"??_);_(@_)', float),
        0x2C: ('_($*#,##0.00_);_($*(#,##0.00);_(*"-"??_);_(@_)', float),
        0x2D: ('mm:ss', datetime.timedelta),
        0x2E: ('[h]:mm:ss', datetime.timedelta),
        0x2F: ('mm:ss.0', datetime.timedelta),
        0x30: ('##0.0E+0', float),
        0x31: ('@', unicode),
    }

    def __init__(self, archive):
        parser = xml.parsers.expat.ParserCreate()
        parser.StartElementHandler = self._start_element
        parser.EndElementHandler = self._end_element
        parser.CharacterDataHandler = self._char_data

        self._numberFormats = []
        self._inCellXfs = False
        self.current_style = None

        fh = archive.styles()
        parser.ParseFile(fh)
        fh.close()

    def _start_element(self, name, attrs):
        if name == 'cellXfs':
            self._inCellXfs = True
        elif self._inCellXfs and name == 'xf':
            self.current_style = {
                'numFmt': int(attrs.get(u'numFmtId', 0)),
                'font': attrs.get(u'fontId'),
                'fill': attrs.get(u'fillId'),
                'border': attrs.get(u'borderId'),
                'xf': attrs.get(u'xfId'),
                'applyFont': attrs.get(u'applyFont'),
                'applyNumFmt': attrs.get(u'applyNumberFormat'),
            }

    def _end_element(self, name):
        if name == 'cellXfs':
            self._inCellXfs = False
        elif self.current_style and name == 'xf':
            self._numberFormats.append(self.current_style)
            self.current_style = None

    def _char_data(self, data):
        pass

    def cell_style(self, styleId):
        """ Return the style with the ID styleId """
        if styleId is None:
            return self._numberFormats[0]   # default format is text '@'
        else:
            return self._numberFormats[int(styleId)]

    def cell_type_from_style(self, style):
        """ Return a python type object from style """
        return self.BUILTIN_NUM_FMTS.get(style['numFmt'], unicode)[self.BUILTIN_TYPE]

    def cell_format_from_style(self, style):
        """ Return a format string for style """
        return self.BUILTIN_NUM_FMTS.get(style['numFmt'], unicode)[self.BUILTIN_FMT]


class Workbook(dict):
    """ Represent the workbook with an index of all work sheets """
    def __init__(self, archive):
        parser = xml.parsers.expat.ParserCreate()
        parser.StartElementHandler = self._start_element
        parser.EndElementHandler = self._end_element
        parser.CharacterDataHandler = self._char_data

        fh = archive.workbook()
        parser.ParseFile(fh)
        fh.close()

    def _start_element(self, name, attrs):
        if name == 'sheet':
            self[attrs['name']] = attrs

    def _end_element(self, name):
        pass

    def _char_data(self, data):
        pass

    def names(self):
        names = self.keys()
        names.sort()
        return names

    def sheet_id(self, name):
        meta = self[name]
        if not meta:
            return
        return meta[u'sheetId']

class Sheet(object):
    """
    Parse a sheet row by row and produce a row_event() for each finished row.
    This is the part where the cell content is parsed too.
    """
    STYLE_IDX = 'i'
    STYLE = 's'
    FMT = 'f'
    REF = 'r'
    COLUMN = 'c'
    VALUE = 'v'
    TYPE = 't'
    TYPE_SHARED_STRING = u's'
    GENERATED_CELL = 'g'

    rel_re = re.compile(r'([A-Z]+)(\d+)')
    MAX_COLUMNS = 1024

    def __init__(self, doc, archive, sheet_id):
        self.document = doc

        parser = xml.parsers.expat.ParserCreate()
        parser.StartElementHandler = self._start_element
        parser.EndElementHandler = self._end_element
        parser.CharacterDataHandler = self._char_data

        self.data = None
        self.is_sheetdata = False
        self.row_count = 0
        self.current_row = None
        self.cell = None
        self.is_value = False

        self.shared_strings = self.document.shared_strings()
        self.styles = self.document.styles()

        fh = archive.sheet(sheet_id)
        parser.ParseFile(fh)
        fh.close()

        del self.shared_strings

    def parse_rel(self, cell):
        """ Convert numeric reference (A1, AD23) into numeric tuple (column, row) """
        column, row = self.rel_re.match(cell[self.REF]).groups()

        v = 0
        for i, ch in enumerate(column):
            s = len(column) - i - 1
            v += (ord(ch) - ord('A') + 1) * (26**s)

        cell[self.REF] = (v, row)
        # savety check to detect omitted cells what we currently don't support
        if cell[self.COLUMN] > v:
            raise Exception(
                "Detected smaller index than current cell, something is wrong! (row %s): %s <> %s" % (
                    row, v, cell[self.COLUMN]
                ))
        # add omitted cells
        for i in xrange(cell[self.COLUMN], v):
            self.current_row.append({
                self.GENERATED_CELL: True,
                self.STYLE_IDX: None,
                self.TYPE: None,
                self.REF: (i, row),
                self.COLUMN: i,
                self.VALUE: u'',
                self.FMT: unicode,
            })

    def _start_element(self, name, attrs):
        #print "start element:", name, attrs
        if name == 'sheetData':
            self.is_sheetdata = True
        elif self.is_sheetdata and name == 'row':
            self.current_row = []
        elif name == 'c':
            self.cell = {
                self.STYLE_IDX: attrs.get(u's'),
                self.TYPE: attrs.get(u't'),
                self.REF: attrs.get(u'r'),
                self.COLUMN: len(self.current_row) + 1,
            }
        elif name == 'v':
            self.is_value = True

    def _end_element(self, name):
        #print "end element:", name
        if name == 'sheetData':
            self.is_sheetdata = False
        elif self.is_sheetdata and name == 'row':
            self.row_count += 1
            self.document.row_event(self.current_row)
            self.current_row = None
        elif name == 'c':
            c = self.cell
            self.parse_rel(c)
            if c[self.TYPE] == self.TYPE_SHARED_STRING:
                idx = int(self.data, 10)
                c[self.VALUE] = self.shared_strings[idx]
            else:
                c[self.VALUE] = self.data
            #fmt = self.styles.numberFormat(c[self.STYLE_IDX])
            #print "cell format is:", str(fmt['numFmt']), c[self.COLUMN], c[self.VALUE]
            c[self.STYLE] = self.styles.cell_style(c[self.STYLE_IDX])
            c[self.FMT] = cellType = self.styles.cell_type_from_style(c[self.STYLE])
            v = c[self.VALUE]
            if (v is not None) and (c[self.FMT] in (datetime.datetime,
                                                    datetime.date,
                                                    datetime.time)):
                try:
                    d = xldate.xldate_as_tuple(float(v), 0)
                    c[self.VALUE] = cellType(*d)
                except (xldate.XLDateAmbiguous, ValueError), e:
                    if v == 1.0:
                        print "value 1.0 for date:", c
                        c[self.VALUE] = ''
                    else:
                        print "Invalid date, assume text or number content:", c[self.VALUE]
                        if re.match(r'^\d+$', v):
                            c[self.TYPE] = cellType = int
                        elif re.match(r'^\d+\.\d+$', v):
                            c[self.TYPE] = cellType = float
                        else:
                            c[self.TYPE] = cellType = unicode
                        c[self.VALUE] = cellType(v)
            else:
                if v is None:
                    c[self.VALUE] = ''
                elif not cellType is unicode:
                    try:
                        c[self.VALUE] = cellType(v)
                    except TypeError, e:
                        print repr(c)
                        print str(e), "value:", repr(v)
                        raise e
            self.current_row.append(c)
            self.data = None
            self.cell = None
        elif name == 'v':
            self.is_value = False

    def _char_data(self, data):
        #print "data value", data
        if self.is_value:
            if self.data is None:
                self.data = data
            else:
                self.data += data

class Document(object):
    """
    Represent a whole XLSX document and provide a high level interface to
    its parts.
    """
    def __init__(self, filename=None):
        self.__archive = None
        self.__shared_strings = None
        self.__styles = None
        self.__workbook = None
        self.__sheets = {}
        self.__row_event_handlers = []
        if filename is not None:
            self.open(filename)

    def open(self, filename):
        """ Provide filename of XLSX document to open """
        self.__archive = DocumentArchive(filename)

    def archive(self):
        if self.__archive is None:
            raise Error("No document specified")
        return self.__archive

    def shared_strings(self):
        if self.__shared_strings is None:
            self.__shared_strings = SharedStrings(self.archive())
        return self.__shared_strings

    def styles(self):
        if self.__styles is None:
            self.__styles = Styles(self.archive())
        return self.__styles

    def workbook(self):
        if self.__workbook is None:
            self.__workbook = Workbook(self.archive())
        return self.__workbook

    def sheet_names(self):
        """ Return the name of the sheets in workbook """
        return self.workbook().names()

    def sheet(self, name):
        """ Return a sheet object from the sheet with the name """
        if not self.__sheets.has_key(name):
            sheet_id = self.workbook().sheet_id(name)
            if not sheet_id:
                return None
            self.__sheets[name] = Sheet(self, self.archive(), sheet_id)
        return self.__sheets[name]

    def add_row_event_handler(self, handler):
        """
        Register a row event handler.
        Such a handler is a function or function object that takes exactly on
        row element as argument: f(row)
        """
        self.__row_event_handlers.append(handler)

    def remove_row_event_handler(self, handler):
        """
        Remove row event handler from registered list of handlers.
        """
        self.__row_event_handlers.remove(handler)

    def row_event(self, row):
        for handler in self.__row_event_handlers:
            handler(row)

class FirstNRowStorage(list):
    """ Stores the first N rows from a worksheet """
    def __init__(self, n=10):
        super(FirstNRowStorage)
        self.n = n
        self.rows = []
        self.is_limit_reached = False

    def __call__(self, row):
        if self.is_limit_reached:
            return
        self.rows.append(row)
        self.is_limit_reached = not (len(self.rows) <= self.n)

def main():
    import sys
    if len(sys.argv) < 2:
        raise Exception("arguments XLSX file or workbook missing")

    doc = Document()
    storage = FirstNRowStorage(2)
    doc.add_row_event_handler(storage)
    doc.open(sys.argv[1])
    print "Read %d shared strings" % len(doc.shared_strings())
    print "Workbook contains sheets:", doc.sheet_names()
    sheetname = sys.argv[2]
    print "Sheet ID for '%s':" % sheetname, doc.workbook().sheet_id(sheetname)
    sheet = doc.sheet(sheetname)
    print "Read %d rows" % len(storage.rows)
    print "row 0:"
    def debug_row(row):
        for c in row:
            print "    %s: %-16s: %s" % (
                c[Sheet.REF],
                doc.styles().cell_format_from_style(c[Sheet.STYLE]),
                c[Sheet.VALUE]
            )
    debug_row(storage.rows[0])
    print "row 1:"
    debug_row(storage.rows[1])

if __name__ == '__main__':
    import cProfile
    cProfile.run('main()', 'profile')
    #main()
