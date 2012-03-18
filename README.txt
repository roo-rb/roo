roo
    by Thomas Preymesser
    http://thopre.wordpress.com

== DESCRIPTION:

Roo can access the contents of various spreadsheet files. It can handle
* Openoffice
* Excel
* Google spreadsheets
* Excelx
* Libreoffice
* CSV


== FEATURES/PROBLEMS:

You don't need to have an installation of Openoffice or Excel. All you need is
the spreadsheet file. It's platform independent so you can read an Excel file
under Linux without having to install the MS office suite.

Roo implements read access for all spreadsheet types and read/write access for
Google spreadsheets.

Unless the underlying 'spreadsheet' gem supports formulas there is no support
for formulas in Roo for .xls files (you get the result of a formula in such a
file but not the formula itself)

== SUPPORT ROO DEVELOPMENT:

If you want to support the further development of the Roo gem, send Bitcoins to <b><code>1KecEuitSFZwx2towBcwbBXmaY5eYjJC9h</code></b>

== SYNOPSIS:

  require 'rubygems'
  require 'roo'

  s = Openoffice.new("myspreadsheet.ods")      # creates an Openoffice Spreadsheet instance
  s = Excel.new("myspreadsheet.xls")           # creates an Excel Spreadsheet instance
  s = Google.new("myspreadsheetkey_at_google") # creates an Google Spreadsheet instance
  s = Excelx.new("myspreadsheet.xlsx")         # creates an Excel Spreadsheet instance for Excel .xlsx files
  s = Csv.new("myspreadsheet.csv")             # creates an Csv Spreadsheet instance for CSV files

  s.default_sheet = s.sheets.first  # first sheet in the spreadsheet file will be used

  # s.sheet is an array which holds the names of the sheets within
  # a spreadsheet.
  # you can also write
  # s.default_sheet = s.sheets[2] or
  # s.default_sheet = 'Sheet 3'or 
  # s.default_sheet = 3                       # please note that sheet numbering
                                              # starts with 1 (not 0) 
  # all these notions above means 'select the third sheet'

  s.cell(1,1)                                 # returns the content of the first row/first cell in the sheet
  s.cell('A',1)                               # same cell
  s.cell(1,'A')                               # same cell
  s.cell(1,'A',s.sheets[0])                   # same cell

  # almost all methods have an optional argument 'sheet'.
  # If this parameter is omitted, the default_sheet will be used.

  s.info                                      # prints infos about the spreadsheet file

  s.first_row                                 # the number of the first row
  s.last_row                                  # the number of the last row
  s.first_column                              # the number of the first column
  s.last_column                               # the number of the last column

  # limited font information is available

  s.font(1,1).bold?
  s.font(1,1).italic?
  s.font(1,1).underline?


see http://roo.rubyforge.org for a more complete tutorial

== REQUIREMENTS:

All dependent gems will be automatically installed.

== INSTALL:

  [sudo] gem install roo

== LICENSE:

(The MIT License)

Copyright (c) 2008-2011 

Permission is hereby granted, free of charge, to any person obtaining
a copy of this software and associated documentation files (the
'Software'), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish,
distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to
the following conditions:

The above copyright notice and this permission notice shall be
included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED 'AS IS', WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
