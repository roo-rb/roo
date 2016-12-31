module Roo
  module Formatters
    module Matrix
      # returns a matrix object from the whole sheet or a rectangular area of a sheet
      def to_matrix(from_row = nil, from_column = nil, to_row = nil, to_column = nil, sheet = default_sheet)
        require 'matrix'

        return ::Matrix.empty unless first_row

        from_row ||= first_row(sheet)
        to_row ||= last_row(sheet)
        from_column ||= first_column(sheet)
        to_column ||= last_column(sheet)

        ::Matrix.rows(from_row.upto(to_row).map do |row|
          from_column.upto(to_column).map do |col|
            cell(row, col, sheet)
          end
        end)
      end
    end
  end
end
