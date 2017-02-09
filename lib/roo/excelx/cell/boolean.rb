module Roo
  class Excelx
    class Cell
      class Boolean < Cell::Base
        attr_reader :value, :formula, :format, :cell_type, :cell_value, :link, :coordinate

        def initialize(value, formula, style, link, coordinate)
          super(value, formula, nil, style, link, coordinate)
          @type = @cell_type = :boolean
          @value = link? ? Roo::Link.new(link, value) : create_boolean(value)
        end

        def formatted_value
          value ? 'TRUE'.freeze : 'FALSE'.freeze
        end

        private

        def create_boolean(value)
          # FIXME: Using a boolean will cause methods like Base#to_csv to fail.
          #       Roo is using some method to ignore false/nil values.
          value.to_i == 1 ? true : false
        end
      end
    end
  end
end
