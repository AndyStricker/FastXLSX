import xml.parsers.expat
import zipfile
import datetime
import re
import xldate

class DocumentArchive(object):
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
    BUILTIN_NUM_FMTS = {
        0: int,
        1: datetime,
        2: datetime,
        3: datetime,
        4: datetime,
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

    def numberFormat(self, styleId):
        style = self._numberFormats[int(styleId)]
        return self.BUILTIN_NUM_FMTS.get(style['numFmt'], unicode)


class Workbook(dict):
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
    STYLE = 's'
    TYPE = 't'
    TYPE_SHARED_STRING = u's'
    REF = 'r'
    COLUMN = 'c'
    VALUE = 'v'

    rel_re = re.compile(r'([A-Z]+)(\d+)')

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

    def _start_element(self, name, attrs):
        #print "start element:", name, attrs
        if name == 'sheetData':
            self.is_sheetdata = True
        elif self.is_sheetdata and name == 'row':
            self.current_row = []
        elif name == 'c':
            self.cell = {
                self.STYLE: attrs.get(u's'),
                self.TYPE: attrs.get(u't'),
                self.REF: attrs.get(u'r'),
                self.COLUMN: len(self.current_row),
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
            if c[self.TYPE] == self.TYPE_SHARED_STRING:
                idx = int(self.data, 10)
                c[self.VALUE] = self.shared_strings[idx]
            else:
                c[self.VALUE] = self.data
            style = self.styles.numberFormat(c[self.STYLE])
            print "cell style is:", str(style)
            if style is datetime and c[self.VALUE] is not None:
                try:
                    v = float(c[self.VALUE])
                    try:
                        d = xldate.xldate_as_tuple(v, 0)
                        c[self.VALUE] = datetime.datetime(*d)
                    except xldate.XLDateAmbiguous, e:
                        if v == 1.0:
                            print "value 1.0 for date:", c
                            c[self.VALUE] = ''
                        else:
                            raise e
                except ValueError, e:
                    print "no float:", c
                    pass
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
        return self.workbook().names()

    def sheet(self, name):
        if not self.__sheets.has_key(name):
            sheet_id = self.workbook().sheet_id(name)
            if not sheet_id:
                return None
            self.__sheets[name] = Sheet(self, self.archive(), sheet_id)
        return self.__sheets[name]

    def add_row_event_handler(self, handler):
        self.__row_event_handlers.append(handler)

    def remove_row_event_handler(self, handler):
        self.__row_event_handlers.remove(handler)

    def row_event(self, row):
        for handler in self.__row_event_handlers:
            handler(row)

def debug_row(row):
    for cell in row:
        print "      [%4s, %c, %4s] %s" % (
            cell['s'],
            cell['type'],
            cell['value']
        )

class FirstNRowStorage(list):
    """ Stores the first N rows from a worksheet """
    def __init__(self, n=10):
        super(FirstNRowStorage)
        self.n = n
        self.rows = []
        self.is_limit = False

    def __call__(self, row):
        if self.is_limit:
            return
        self.rows.append(row)
        self.is_limit = not (len(self.rows) <= self.n)

def main():
    import sys
    doc = Document()
    doc.open(sys.argv[1])
    storage = FirstNRowStorage(2)
    doc.add_row_event_handler(storage)
    print "Read %d shared strings" % len(doc.shared_strings())
    print "Workbook contains sheets:", doc.sheet_names()
    #sheetname = u'Bild'
    sheetname = u'Personendaten'
    print "Sheet ID for '%s':" % sheetname, doc.workbook().sheet_id(sheetname)
    sheet = doc.sheet(sheetname)
    print "Read %d rows" % sheet.row_count
    print "row 0:"
    debug_row(storage.rows[0])
    print "row 1:"
    debug_row(storage.rows[1])

if __name__ == '__main__':
    import cProfile
    cProfile.run('main()', 'profile')
    #main()
