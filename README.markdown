# README for Roo

Roo implements read access for all spreadsheet types and read/write access for
Google spreadsheets. It can handle
* Openoffice
* Excel
* Google spreadsheets
* Excelx
* Libreoffice
* CSV

Using Roo to access Google spreadsheets requires you install the 'google-spreadsheet-ruby' gem separately.

Unless the underlying 'spreadsheet' gem supports formulas there is no support
for formulas in Roo for .xls files (you get the result of a formula in such a
file but not the formula itself)

## Usage:

    require 'roo'

    s = Roo::Openoffice.new("myspreadsheet.ods")      # creates an Openoffice Spreadsheet instance
    s = Roo::Excel.new("myspreadsheet.xls")           # creates an Excel Spreadsheet instance
    s = Roo::Google.new("myspreadsheetkey_at_google") # creates an Google Spreadsheet instance
    s = Roo::Excelx.new("myspreadsheet.xlsx")         # creates an Excel Spreadsheet instance for Excel .xlsx files

    s.default_sheet = s.sheets.first             # first sheet in the spreadsheet file will be used

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

# Fork Changelog / New Features

    # Spreadsheet.open can accept both files and paths

    xls = Roo::Spreadsheet.open('./new_prices.xls')

    # no more setting xls.default_sheet, just use this

    xls.sheet('Info').row_count
    xls.sheet(0).row_count

    # excel likes to create random "Data01" sheets for macros
    # use this to find the sheet with the most data to parse

    xls.longest_sheet

    # this excel file has multiple worksheets, let's iterate through each of them and process

    xls.each_with_pagename do |name,sheet|
    puts sheet.row_count
    end

    # pull out a hash of exclusive column data (get rid of useless columns and save memory)

    xls.each(:id => 'UPC',:qty => 'ATS') {|hash| arr << hash}
    #=> hash will appear like {:upc=>727880013358, :qty => 12}

    # NOTE: .parse does the same as .each, except it returns an array (similar to each vs. map)

    # not sure exactly what a column will be named? try a wildcard search with the character *
    # regex characters are allowed ('^price\s')
    # case insensitive

    xls.parse(:id => 'UPC*SKU',:qty => 'ATS*\sATP\s*QTY$')

    # if you need to locate the header row and assign the header names themselves,
    # use the :header_search option

    xls.parse(:header_search => ['UPC*SKU','ATS*\sATP\s*QTY$'])
    #=> each element will appear in this fashion:
    #=> {"UPC" => 123456789012, "STYLE" => "987B0", "COLOR" => "blue", "QTY" => 78}

    # want to strip out annoying unicode characters and surrounding white space?

    xls.parse(:clean => true)

    # another bonus feature is a patch to prevent the Spreadsheet gem from parsing
    # thousands and thousands of blank lines. i got fed up after watching my computer
    # nearly catch fire for 4 hours for a spreadsheet with only 200 ACTUAL lines
    # - located in lib/roo/worksheet.rb

