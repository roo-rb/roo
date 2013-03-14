=begin
require 'rubygems'
require 'roo'

oo = Excel.new("tmp.xls")
oo.default_sheet = oo.sheets.first
oo.first_row.upto(oo.last_row) do |row|
	oo.first_column.upto(oo.last_column) do |col|
		p oo.cell(row,col)
	end
end
FileUtils.rm_f("tmp.xls", {:verbose => true, :force => true})
=end
require 'spreadsheet'

book = Spreadsheet.open 'tmp.xls'
sheet = book.worksheet 0
sheet.each do |row| puts row[0] end

FileUtils.rm("tmp.xls")
