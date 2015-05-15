require 'date'

module Roo
  class Excelx
    class Cell
      attr_reader :type, :formula, :value, :excelx_type, :excelx_value, :style, :hyperlink, :coordinate
      attr_writer :value

      def initialize(value, type, formula, excelx_type, excelx_value, style, hyperlink, base_date, coordinate)
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
        if @formula
          :formula
        elsif @value.is_a?(Roo::Link)
          :link
        else
          @type
        end
      end

      class Coordinate
        attr_accessor :row, :column

        def initialize(row, column)
          @row, @column = row, column
        end
      end

      private

      def type_cast_value(value)
        case @type
        when :float, :percentage
          value.to_f
        when :date
          yyyy, mm, dd = (@base_date + value.to_i).strftime('%Y-%m-%d').split('-')
          Date.new(yyyy.to_i, mm.to_i, dd.to_i)
        when :datetime
          create_datetime_from((@base_date + value.to_f.round(6)).strftime('%Y-%m-%d %H:%M:%S.%N'))
        when :time
          value.to_f * 86_400
        when :string
          value
        else
          value
        end
      end

      def create_datetime_from(datetime_string)
        date_part, time_part = round_time_from(datetime_string).split(' ')
        yyyy, mm, dd = date_part.split('-')
        hh, mi, ss = time_part.split(':')
        DateTime.civil(yyyy.to_i, mm.to_i, dd.to_i, hh.to_i, mi.to_i, ss.to_i)
      end

      def round_time_from(datetime_string)
        date_part, time_part = datetime_string.split(' ')
        yyyy, mm, dd = date_part.split('-')
        hh, mi, ss = time_part.split(':')

        Time.new(yyyy.to_i, mm.to_i, dd.to_i, hh.to_i, mi.to_i, ss.to_r).round(0).strftime('%Y-%m-%d %H:%M:%S')
      end
    end
  end
end
