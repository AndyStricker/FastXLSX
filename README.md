Reader for XLSX files with a minimum memory footprint
=====================================================

Don't use this if you just want to read a XLSX file. There are a few other
projects that will do that better for now. This piece of code is optimized to
read very large files.

Before I started this project I tried the following projects. All of them have
inspired me. I'll like to give credit to the authors of those projects. While I
also consulted the ECMA-376 specifications [1], the source code of those
projects were easier to understand.

Project:    openpyxl
Author:     Eric Gazoni
Homepage:   http://ericgazoni.wordpress.com/2010/04/10/openpyxl-python-xlsx/
Repository: http://bitbucket.org/ericgazoni/openpyxl/

Project:    python-xlsx
Author:     Ståle Undheim
Repository: https://github.com/staale/python-xlsx

Project:    pyXLSX
Author:     Lee Gao
Repository: https://github.com/leegao/pyXLSX

Project:    Spreadsheet::XLSX
Author:     Dmitry Ovsyanko
CPAN:       http://search.cpan.org/~dmow/Spreadsheet-XLSX-0.13-withoutworldwriteables/lib/Spreadsheet/XLSX.pm

Usage
-----

Maybe you need to set PYTHONPATH environment variable to find the fastxslx
package first.

python xlsx-to-csv.py YOUR_EXCEL_FILE.xslx YOUR_WORKSHEET_NAME

The converter then will create a file named as follows:

  YOUR_EXCEL_FILE.xslx-YOUR_WORKSHEET_NAME.csv

Todo
----

- Support custom formats in cells

History
-------

This project started as I got a database dump as a XLSX file instead something
more portable. As this is usually not really a problem - just open it with
OpenOffice/LibreOffice and save it as CSV - the huge size of one of those files
set a limit: It was about 400 MB in size and contained a table with more than a
million rows. LibreOffice didn't choke on them, but produced an empty table
where the row count passes a million rows.

So I tried the four listed projects above but failed on their memory
consumption: All of them used a DOM or Tree/Node based XML parser, that
resulted with holding about twice the size of the file in memory. (Note that a
400 MB XLSX file is the compressed size, the memory usage will be much higher)

So I started to parse it by my own, using the simplest and fastest event based
XML parser I know of: expat. The document is parsed row by row, each time
calling a row event handler. It only needs to keep the shared string table in
memory. Even the document ZIP container is read as a stream. I successful use
this parser to convert hugh database export tables to CSV, the biggest document
was 400 MB is size.


[1] http://www.ecma-international.org/publications/standards/Ecma-376.htm
