require 'date'
require 'roo/excelx/cell/base'
require 'roo/excelx/cell/boolean'
require 'roo/excelx/cell/datetime'
require 'roo/excelx/cell/date'
require 'roo/excelx/cell/empty'
require 'roo/excelx/cell/number'
require 'roo/excelx/cell/string'
require 'roo/excelx/cell/time'

module Roo
  class Excelx
    class Cell
      attr_reader :formula, :value, :excelx_type, :excelx_value, :style, :hyperlink, :coordinate
      attr_writer :value

      # DEPRECATED: Please use Cell.create_cell instead.
      def initialize(value, type, formula, excelx_type, excelx_value, style, hyperlink, base_date, coordinate)
        warn '[DEPRECATION] `Cell.new` is deprecated.  Please use `Cell.create_cell` instead.'
        @type = type
        @formula = formula
        @base_date = base_date if [:date, :datetime].include?(@type)
        @excelx_type = excelx_type
        @excelx_value = excelx_value
        @style = style
        @value = type_cast_value(value)
        @value = Roo::Link.new(hyperlink, @value.to_s) if hyperlink
        @coordinate = coordinate
      end

      def type
        case
        when @formula
          :formula
        when @value.is_a?(Roo::Link)
          :link
        else
          @type
        end
      end

      def self.create_cell(type, *values)
        cell_class(type)&.new(*values)
      end

      def self.cell_class(type)
        case type
        when :string
          Cell::String
        when :boolean
          Cell::Boolean
        when :number
          Cell::Number
        when :date
          Cell::Date
        when :datetime
          Cell::DateTime
        when :time
          Cell::Time
        end
      end

      # Deprecated: use Roo::Excelx::Coordinate instead.
      class Coordinate
        attr_accessor :row, :column

        def initialize(row, column)
          warn '[DEPRECATION] `Roo::Excel::Cell::Coordinate` is deprecated.  Please use `Roo::Excelx::Coordinate` instead.'
          @row, @column = row, column
        end
      end

      private

      def type_cast_value(value)
        case @type
        when :float, :percentage
          value.to_f
        when :date
          create_date(@base_date + value.to_i)
        when :datetime
          create_datetime(@base_date + value.to_f.round(6))
        when :time
          value.to_f * 86_400
        else
          value
        end
      end

      def create_date(date)
        yyyy, mm, dd = date.strftime('%Y-%m-%d').split('-')

        ::Date.new(yyyy.to_i, mm.to_i, dd.to_i)
      end

      def create_datetime(date)
        datetime_string = date.strftime('%Y-%m-%d %H:%M:%S.%N')
        t = round_datetime(datetime_string)

        ::DateTime.civil(t.year, t.month, t.day, t.hour, t.min, t.sec)
      end

      def round_datetime(datetime_string)
        /(?<yyyy>\d+)-(?<mm>\d+)-(?<dd>\d+) (?<hh>\d+):(?<mi>\d+):(?<ss>\d+.\d+)/ =~ datetime_string

        ::Time.new(yyyy.to_i, mm.to_i, dd.to_i, hh.to_i, mi.to_i, ss.to_r).round(0)
      end
    end
  end
end
