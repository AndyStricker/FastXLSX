import sys
import csv
import datetime
import fastxlsx

class Converter(object):
    VALUE = fastxlsx.reader.Sheet.VALUE

    def __init__(self, outfile, with_progress=False):
        if isinstance(outfile, str):
            self.outfile = file(outfile, 'w')
        else:
            self.outfile = outfile
        self.writer = csv.writer(self.outfile,
                                 delimiter=',',
                                 quotechar='"',
                                 quoting=csv.QUOTE_NONNUMERIC)
        self.with_progress = with_progress
        self.first_row = None
        self.columns = 0
        self.rows = 0

    def __call__(self, row):
        record = []
        for cell in row:
            v = cell[self.VALUE]
            if v is None:
                record.append('')
            elif isinstance(v, datetime.datetime):
                record.append(v.strftime('%d.%m.%Y %H:%M:%S'))
            else:
                try:
                    v = int(v, 10)
                except ValueError:
                    try:
                        v = float(v)
                    except ValueError:
                        v = v.encode('UTF-8')
                record.append(v)

        if self.first_row is None:
            self.columns = len(record)
            self.first_row = record
            print self.first_row
        else:
            record.extend(['' for x in xrange(len(record), self.columns)])

        self.writer.writerow(record)
        self.rows += 1
        if self.with_progress:
            self.update_progress()

    def close(self):
        if self.outfile:
            self.outfile.close()
            self.outfile = None

    def update_progress(self):
        r = self.rows
        if (r % 64) == 0:
            print "\rRow %d       " % r,
            sys.stdout.flush()

