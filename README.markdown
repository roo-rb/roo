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

# README for Roo

Roo is available here and on Rubyforge. You can install the official release with 'gem install roo' or refer to the installation instructions below for the latest development gem. 

NOTE: Roo 1.9 was released by Thomas and I think it was intended for Ruby 1.9 but the dependencies are not working properly so everyone gets it with a gem install Roo. I'm trying to get a hold of him to work out how to fix things.

In the meantime, Roo 1.3.11 should be on gemcutter and works with Ruby 1.8 with no known issues. I'll continue to maintain this version in the interim.  

## Installation

    # Run the following if you haven't done so before:
    gem sources -a http://gems.github.com/

    # Install the gem:
    sudo gem install roo -v 1.3.11

## Usage:

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

