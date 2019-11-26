
module Roo
  class Excelx
    class Cell
      class Empty < Cell::Base
        attr_reader :value, :formula, :format, :cell_type, :cell_value, :coordinate

        attr_reader_with_default default_type: nil, style: nil

        def initialize(coordinate)
          @coordinate = coordinate
        end

        def empty?
          true
        end
      end
    end
  end
end
