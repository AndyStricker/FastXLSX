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

import sys
import csv
import datetime
import fastxlsx

class Converter(object):
    """ Convert XSLX to CSV. (with progress status display) """
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
            elif isinstance(v, (unicode, str)):
                record.append(v.encode('UTF-8'))
            elif isinstance(v, (float, int)):
                record.append(str(v))
            elif isinstance(v, datetime.datetime):
                record.append(v.strftime('%Y-%m-%dT%H:%M:%S'))
            elif isinstance(v, datetime.date):
                record.append(v.strftime('%Y-%m-%d'))
            elif isinstance(v, datetime.time):
                record.append(v.strftime('%H:%M:%S'))
            else:
                raise Exception("Unknown format detected: " + repr(v))

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

