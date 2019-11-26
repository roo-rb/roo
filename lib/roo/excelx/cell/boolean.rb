# frozen_string_literal: true

module Roo
  class Excelx
    class Cell
      class Boolean < Cell::Base
        attr_reader :value, :formula, :format, :cell_value, :coordinate

        attr_reader_with_default default_type: :boolean, cell_type: :boolean

        def initialize(value, formula, style, link, coordinate)
          super(value, formula, nil, style, nil, coordinate)
          @value = link ? Roo::Link.new(link, value) : create_boolean(value)
        end

        def formatted_value
          value ? 'TRUE' : 'FALSE'
        end

        private

        def create_boolean(value)
          # FIXME: Using a boolean will cause methods like Base#to_csv to fail.
          #       Roo is using some method to ignore false/nil values.
          value.to_i == 1
        end
      end
    end
  end
end
