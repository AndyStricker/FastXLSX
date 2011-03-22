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

filename = sys.argv[1]
if not filename:
    print "No filename specified"
    sys.exit(1)
sheetname = sys.argv[2]
if not sheetname:
    print "No sheetname specified"

outfile = file("%s-%s.csv" % (filename, sheetname), 'w')

doc = fastxlsx.reader.Document()
print "Loading workbook in progress..."
print "    - create CSV converter"
handler = fastxlsx.csvconverter.Converter(outfile, with_progress=True)
print "    - open input file %s ..." % filename
doc.add_row_event_handler(handler)
doc.open(sys.argv[1])
print "    - shared strings...",
count = len(doc.shared_strings())
print " Loaded %d shared strings" % count
print "    - Workbook contains sheets:", doc.sheet_names()
print "    - select work sheet %s" % sheetname
sheet = doc.sheet(sheetname)
if not sheet:
    print "No such sheet found:", sheetname
    sys.exit(1)

print "\nRead %d rows with %d columns" % (handler.rows, handler.columns)
print "Done converting document"
outfile.close()
