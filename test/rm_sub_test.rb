# Loeschen von Dateien, wenn mit Excel geoeffnet
require 'rubygems'
require 'roo'

oo = Excel.new("tmp.xls")
#oo = Openoffice.new("tmp.ods")
oo.default_sheet = oo.sheets.first
oo.first_row.upto(oo.last_row) do |row|
	oo.first_column.upto(oo.last_column) do |col|
		p oo.cell(row,col)
	end
end

