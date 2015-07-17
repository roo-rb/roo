module Roo
  class Excelx
    class Cell
      class String < Cell::Base
        attr_reader :value, :formula, :format, :cell_type, :cell_value, :link, :coordinate

        def initialize(value, formula, style, link, coordinate)
          super(value, formula, nil, style, link, coordinate)
          @type = @cell_type = :string
          @value = link? ? Roo::Link.new(link, value) : value
        end

        def empty?
          value.empty?
        end
      end
    end
  end
end
