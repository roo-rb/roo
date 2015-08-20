module Roo
  class Excelx
    class Coordinate
      attr_accessor :row, :column

      def initialize(row, column)
        @row = row
        @column = column
      end
    end
  end
end
