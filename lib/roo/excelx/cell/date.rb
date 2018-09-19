require 'date'

module Roo
  class Excelx
    class Cell
      class Date < Roo::Excelx::Cell::DateTime
        attr_reader :value, :formula, :format, :cell_type, :cell_value, :coordinate

        attr_reader_with_default default_type: :date

        def initialize(value, formula, excelx_type, style, link, base_date, coordinate)
          # NOTE: Pass all arguments to the parent class, DateTime.
          super
          @format = excelx_type.last
          @value = link ? Roo::Link.new(link, value) : create_date(base_date, value)
        end

        private

        def create_datetime(_,_);  end

        def create_date(base_date, value)
          base_date + value.to_i
        end
      end
    end
  end
end
