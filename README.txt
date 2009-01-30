README for roo
==============

Installation:

    sudo gem install roo

Usage:

    require 'rubygems'
    require 'roo'

    s = Openoffice.new("myspreadsheet.ods")      # creates an Openoffice Spreadsheet instance
    s = Excel.new("myspreadsheet.xls")           # creates an Excel Spreadsheet instance
    s = Google.new("myspreadsheetkey_at_google") # creates an Google Spreadsheet instance
    s = Excelx.new("myspreadsheet.xlsx")         # creates an Excel Spreadsheet instance for Excel .xlsx files

    s.default_sheet = s.sheets.first  # first sheet in the spreadsheet file will be used

    # s.sheet is an array which holds the names of the sheets within
    # a spreadsheet.
    # you can also write
    # s.default_sheet = s.sheets[3] or
    # s.default_sheet = 'Sheet 3'

    s.cell(1,1)                                 # returns the content of the first row/first cell in the sheet
    s.cell('A',1)                               # same cell
    s.cell(1,'A')                               # same cell
    s.cell(1,'A',s.sheets[0])                   # same cell

    # almost all methods have an optional argument 'sheet'.
    # If this parameter is ommitted, the default_sheet will be used.
    
    s.info                                      # prints infos about the spreadsheet file

    s.first_row                                 # the number of the first row
    s.last_row                                  # the number of the last row
    s.first_column                              # the number of the first column
    s.last_column                               # the number of the last column


see http://roo.rubyforge.org for a more complete tutorial

