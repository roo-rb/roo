module Roo
  class Excelx
    class Coordinate
      attr_accessor :row, :column

      def initialize(row, column)
        @row = row
        @column = column
      end

      def to_a
        @array ||= [row, column].freeze
      end
    end
  end
end
