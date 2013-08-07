require 'spreadsheet'

book = Spreadsheet.open 'tmp.xls'
sheet = book.worksheet 0
sheet.each do |row| puts row[0] end

FileUtils.rm("tmp.xls")
