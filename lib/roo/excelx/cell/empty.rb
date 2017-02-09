
module Roo
  class Excelx
    class Cell
      class Empty < Cell::Base
        attr_reader :value, :formula, :format, :cell_type, :cell_value, :hyperlink, :coordinate

        def initialize(coordinate)
          @value = @formula = @format = @cell_type = @cell_value = @hyperlink = nil
          @coordinate = coordinate
        end

        def empty?
          true
        end
      end
    end
  end
end
