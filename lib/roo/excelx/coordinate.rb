module Roo
  class Excelx
    class Coordinate < ::Array

      def initialize(row, column)
        super() << row << column
        freeze
      end

      def row
        self[0]
      end

      def column
        self[1]
      end
    end
  end
end
