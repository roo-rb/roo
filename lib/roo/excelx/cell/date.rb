require 'date'

module Roo
  class Excelx
    class Cell
      class Date < Roo::Excelx::Cell::DateTime
        attr_reader :value, :formula, :format, :cell_type, :cell_value, :link, :coordinate

        def initialize(value, formula, excelx_type, style, link, base_date, coordinate)
          # NOTE: Pass all arguments to the parent class, DateTime.
          super
          @type = :date
          @format = excelx_type.last
          @value = link? ? Roo::Link.new(link, value) : create_date(base_date, value)
        end

        private

        def create_date(base_date, value)
          date = base_date + value.to_i
          yyyy, mm, dd = date.strftime('%Y-%m-%d').split('-')

          ::Date.new(yyyy.to_i, mm.to_i, dd.to_i)
        end
      end
    end
  end
end
