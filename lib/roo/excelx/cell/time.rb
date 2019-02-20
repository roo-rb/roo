require 'date'

module Roo
  class Excelx
    class Cell
      class Time < Roo::Excelx::Cell::DateTime
        attr_reader :value, :formula, :format, :cell_value, :coordinate

        attr_reader_with_default default_type: :time

        def initialize(value, formula, excelx_type, style, link, base_date, coordinate)
          # NOTE: Pass all arguments to DateTime super class.
          super
          @format = excelx_type.last
          @datetime = create_datetime(base_date, value)
          @value = link ? Roo::Link.new(link, value) : (value.to_f * 86_400).to_i
        end

        def formatted_value
          formatter = @format.gsub(/#{TIME_FORMATS.keys.join('|')}/, TIME_FORMATS)
          @datetime.strftime(formatter)
        end

        alias_method :to_s, :formatted_value

        private

        # def create_datetime(base_date, value)
        #   date = base_date + value.to_f.round(6)
        #   datetime_string = date.strftime('%Y-%m-%d %H:%M:%S.%N')
        #   t = round_datetime(datetime_string)
        #
        #   ::DateTime.civil(t.year, t.month, t.day, t.hour, t.min, t.sec)
        # end

        # def round_datetime(datetime_string)
        #   /(?<yyyy>\d+)-(?<mm>\d+)-(?<dd>\d+) (?<hh>\d+):(?<mi>\d+):(?<ss>\d+.\d+)/ =~ datetime_string
        #
        #   ::Time.new(yyyy.to_i, mm.to_i, dd.to_i, hh.to_i, mi.to_i, ss.to_r).round(0)
        # end
      end
    end
  end
end
