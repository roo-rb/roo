require 'forwardable'
module Roo
  class Excelx
    class Sheet
      extend Forwardable

      delegate [:styles, :workbook, :shared_strings, :rels_files, :sheet_files, :comments_files, :image_rels] => :@shared

      attr_reader :images

      def initialize(name, shared, sheet_index, options = {})
        @name = name
        @shared = shared
        @sheet_index = sheet_index
        @images = Images.new(image_rels[sheet_index]).list
        @rels = Relationships.new(rels_files[sheet_index])
        @comments = Comments.new(comments_files[sheet_index])
        @sheet = SheetDoc.new(sheet_files[sheet_index], @rels, shared, options)
      end

      def cells
        @cells ||= @sheet.cells(@rels)
      end

      def present_cells
        @present_cells ||= begin
          warn %{
[DEPRECATION] present_cells is deprecated. Alternate:
  with activesupport    => cells[key].presence
  without activesupport => cells[key]&.presence
          }
          cells.select { |_, cell| cell&.presence }
        end
      end

      # Yield each row as array of Excelx::Cell objects
      # accepts options max_rows (int) (offset by 1 for header),
      # pad_cells (boolean) and offset (int)
      def each_row(options = {}, &block)
        row_count = 0
        options[:offset] ||= 0
        @sheet.each_row_streaming do |row|
          break if options[:max_rows] && row_count == options[:max_rows] + options[:offset] + 1
          if block_given? && !(options[:offset] && row_count < options[:offset])
            block.call(cells_for_row_element(row, options))
          end
          row_count += 1
        end
      end

      def row(row_number)
        first_column.upto(last_column).map do |col|
          cells[[row_number, col]]&.value
        end
      end

      def column(col_number)
        first_row.upto(last_row).map do |row|
          cells[[row, col_number]]&.value
        end
      end

      # returns the number of the first non-empty row
      def first_row
        @first_row ||= first_last_row_col[:first_row]
      end

      def last_row
        @last_row ||= first_last_row_col[:last_row]
      end

      # returns the number of the first non-empty column
      def first_column
        @first_column ||= first_last_row_col[:first_column]
      end

      # returns the number of the last non-empty column
      def last_column
        @last_column ||= first_last_row_col[:last_column]
      end

      def excelx_format(key)
        cell = cells[key]
        styles.style_format(cell.style).to_s if cell
      end

      def hyperlinks
        @hyperlinks ||= @sheet.hyperlinks(@rels)
      end

      def comments
        @comments.comments
      end

      def dimensions
        @sheet.dimensions
      end

      private

      # Take an xml row and return an array of Excelx::Cell objects
      # optionally pad array to header width(assumed 1st row).
      # takes option pad_cells (boolean) defaults false
      def cells_for_row_element(row_element, options = {})
        return [] unless row_element
        cell_col = 0
        cells = []
        @sheet.each_cell(row_element) do |cell|
          cells.concat(pad_cells(cell, cell_col)) if options[:pad_cells]
          cells << cell
          cell_col = cell.coordinate.column
        end
        cells
      end

      def pad_cells(cell, last_column)
        pad = []
        (cell.coordinate.column - 1 - last_column).times { pad << nil }
        pad
      end

      def first_last_row_col
        @first_last_row_col ||= begin
          first_row = last_row = first_col = last_col = nil

          cells.each do |(row, col), cell|
            next unless cell&.presence
            first_row ||= row
            last_row ||= row
            first_col ||= col
            last_col ||= col

            if row > last_row
              last_row = row
            elsif row < first_row
              first_row = row
            end

            if col > last_col
              last_col = col
            elsif col < first_col
              first_col = col
            end
          end

          {first_row: first_row, last_row: last_row, first_column: first_col, last_column: last_col}
        end
      end
    end
  end
end
