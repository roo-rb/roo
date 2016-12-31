# returns an XML representation of all sheets of a spreadsheet file
module Roo
  module Formatters
    module XML
      def to_xml
        Nokogiri::XML::Builder.new do |xml|
          xml.spreadsheet do
            sheets.each do |sheet|
              self.default_sheet = sheet
              xml.sheet(name: sheet) do |x|
                if first_row && last_row && first_column && last_column
                  # sonst gibt es Fehler bei leeren Blaettern
                  first_row.upto(last_row) do |row|
                    first_column.upto(last_column) do |col|
                      next if empty?(row, col)

                      x.cell(cell(row, col),
                      row: row,
                      column: col,
                      type: celltype(row, col))
                    end
                  end
                end
              end
            end
          end
        end.to_xml
      end
    end
  end
end
