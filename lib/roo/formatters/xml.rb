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

                      attributes = { row: row, column: col, type: celltype(row, col) }

                      font = font(row, col)
                      attributes.merge!(bold: font.bold?, italic: font.italic?, underline: font.underline?) if font

                      fill = fills(row, col)
                      attributes.merge!(cell_color: fill.color) if fill

                      x.cell(cell(row, col), attributes)
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
