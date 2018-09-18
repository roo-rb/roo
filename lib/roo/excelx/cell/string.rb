module Roo
  class Excelx
    class Cell
      class String < Cell::Base
        attr_reader :value, :formula, :format, :cell_value, :coordinate

        attr_reader_with_default default_type: :string, cell_type: :string

        def initialize(value, formula, style, link, coordinate)
          super(value, formula, nil, style, link, coordinate)
        end

        def empty?
          value.empty?
        end
      end
    end
  end
end
