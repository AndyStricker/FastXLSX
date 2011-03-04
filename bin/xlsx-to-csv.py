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
handler = fastxlsx.csv.Converter(outfile, with_progress=True)
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
