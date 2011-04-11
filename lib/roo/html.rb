module Roo
  class Spreadsheet
    class HTML < GenericSpreadsheet
  
      def initialize(filename)
        raw = File.open(filename).read
        table = raw.match(/<table>(.*)<\/table>/)[1] rescue raise('No <table> found.')
        @sheets = [(@default_sheet = 'Table')]
        @cell = {@default_sheet => {}}
        parse_html(table)
        @cells_read = {@default_sheet => true}
        @first_row = Hash.new
        @last_row = Hash.new
        @first_column = Hash.new
        @last_column = Hash.new
        @header_line = 1
      end

      def parse_html(raw_string)
        doc = Nokogiri::HTML(raw_string)
        doc.xpath('//tr').each_with_index do |row,row_no|
          row.children.each_with_index do |col,col_no|
            @cell[@default_sheet][[row_no + 1,col_no + 1]] = col.text
          end
        end
      end

      def row(index)
        @cell[@default_sheet].select {|k,v| k[0] == index}.values
      end
  
      def cell(row_no,col_no)
        @cell[@default_sheet][[row_no,col_no]]
      end
      
    end
  end
end