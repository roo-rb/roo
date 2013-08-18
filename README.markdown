# README for Roo

Roo implements read access for all spreadsheet types and read/write access for
Google spreadsheets. It can handle
* OpenOffice
* Excel
* Google spreadsheets
* Excelx
* LibreOffice
* CSV

## Notes

### XLS

There is no support for formulas in Roo for .xls files - you can get the result
of a formula but not the formula itself.

### Google Spreadsheet

Using Roo to access Google spreadsheets requires you install the 'google-spreadsheet-ruby' gem separately.

## License

While Roo is licensed under the MIT / Expat license, please note that the 'spreadsheet' gem [is released under](https://github.com/zdavatz/spreadsheet/blob/master/LICENSE.txt) the GPLv3 license.

## Usage:

```ruby
require 'roo'

s = Roo::OpenOffice.new("myspreadsheet.ods")      # loads an OpenOffice Spreadsheet
s = Roo::Excel.new("myspreadsheet.xls")           # loads an Excel Spreadsheet
s = Roo::Google.new("myspreadsheetkey_at_google") # loads a Google Spreadsheet
s = Roo::Excelx.new("myspreadsheet.xlsx")         # loads an Excel Spreadsheet for Excel .xlsx files
s = Roo::CSV.new("mycsv.csv")                     # loads a CSV file

# You can use CSV to load TSV files, or files of a certain encoding by passing
# in options under the :csv_options key
s = Roo::CSV.new("mytsv.tsv", csv_options: {col_sep: "\t"}) # TSV
s = Roo::CSV.new("mycsv.csv", csv_options: {encoding: Encoding::ISO_8859_1}) # csv with explicit encoding

s.default_sheet = s.sheets.first             # first sheet in the spreadsheet file will be used

# s.sheets is an array which holds the names of the sheets within
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


# Spreadsheet.open can accept both files and paths

xls = Roo::Spreadsheet.open('./new_prices.xls')

# If the File.path or provided path string does not have an extension, you can optionally
# provide one as a string or symbol

xls = Roo::Spreadsheet.open('./rails_temp_upload', extension: :xls)

# no more setting xls.default_sheet, just use this

xls.sheet('Info').row(1)
xls.sheet(0).row(1)

# excel likes to create random "Data01" sheets for macros
# use this to find the sheet with the most data to parse

xls.longest_sheet

# this excel file has multiple worksheets, let's iterate through each of them and process

xls.each_with_pagename do |name, sheet|
  p sheet.row(1)
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
```
