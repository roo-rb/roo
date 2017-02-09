require 'date'

module Roo
  class Excelx
    class Cell
      class DateTime < Cell::Base
        attr_reader :value, :formula, :format, :cell_value, :link, :coordinate

        def initialize(value, formula, excelx_type, style, link, base_date, coordinate)
          super(value, formula, excelx_type, style, link, coordinate)
          @type = :datetime
          @format = excelx_type.last
          @value = link? ? Roo::Link.new(link, value) : create_datetime(base_date, value)
        end

        # Public: Returns formatted value for a datetime. Format's can be an
        #         standard excel format, or a custom format.
        #
        #         Standard formats follow certain conventions. Date fields for
        #         days, months, and years are separated with hyhens or
        #         slashes ("-", /") (e.g. 01-JAN, 1/13/15). Time fields for
        #         hours, minutes, and seconds are separated with a colon (e.g.
        #         12:45:01).
        #
        #         If a custom format follows those conventions, then the custom
        #         format will be used for the a cell's formatted value.
        #         Otherwise, the formatted value will be in the following
        #         format: 'YYYY-mm-dd HH:MM:SS' (e.g. "2015-07-10 20:33:15").
        #
        # Examples
        #    formatted_value #=> '01-JAN'
        #
        # Returns a String representation of a cell's value.
        def formatted_value
          formatter = @format.downcase.split(' ').map do |part|
            if (parsed_format = parse_date_or_time_format(part))
              parsed_format
            else
              warn 'Unable to parse custom format. Using "YYYY-mm-dd HH:MM:SS" format.'
              return @value.strftime('%F %T')
            end
          end.join(' ')

          @value.strftime(formatter)
        end

        private

        def parse_date_or_time_format(part)
          date_regex = /(?<date>[dmy]+[\-\/][dmy]+([\-\/][dmy]+)?)/
          time_regex = /(?<time>(\[?[h]\]?+:)?[m]+(:?ss|:?s)?)/

          if part[date_regex] == part
            formats = DATE_FORMATS
          elsif part[time_regex]
            formats = TIME_FORMATS
          else
            return false
          end

          part.gsub(/#{formats.keys.join('|')}/, formats)
        end

        DATE_FORMATS = {
          'yyyy' => '%Y',  # Year: 2000
          'yy' => '%y',    # Year: 00
          # mmmmm => J-D
          'mmmm' => '%B',  # Month: January
          'mmm' => '%^b',   # Month: JAN
          'mm' => '%m',    # Month: 01
          'm' => '%-m',    # Month: 1
          'dddd' => '%A',  # Day of the Week: Sunday
          'ddd' => '%^a',   # Day of the Week: SUN
          'dd' => '%d',    # Day of the Month: 01
          'd' => '%-d'    # Day of the Month: 1
          # '\\\\'.freeze => ''.freeze,  # NOTE: Fixes a custom format's output.
        }

        TIME_FORMATS = {
          'hh' => '%H',    # Hour (24): 01
          'h' => '%-k'.freeze,    # Hour (24): 1
          # 'hh'.freeze => '%I'.freeze,    # Hour (12): 08
          # 'h'.freeze => '%-l'.freeze,    # Hour (12): 8
          'mm' => '%M',    # Minute: 01
          # FIXME: is this used? Seems like 'm' is used for month, not minute.
          'm' => '%-M',    # Minute: 1
          'ss' => '%S',    # Seconds: 01
          's' => '%-S',    # Seconds: 1
          'am/pm' => '%p', # Meridian: AM
          '000' => '%3N',  # Fractional Seconds: thousandth.
          '00' => '%2N',   # Fractional Seconds: hundredth.
          '0' => '%1N'    # Fractional Seconds: tenths.
        }

        def create_datetime(base_date, value)
          date = base_date + value.to_f.round(6)
          datetime_string = date.strftime('%Y-%m-%d %H:%M:%S.%N')
          t = round_datetime(datetime_string)

          ::DateTime.civil(t.year, t.month, t.day, t.hour, t.min, t.sec)
        end

        def round_datetime(datetime_string)
          /(?<yyyy>\d+)-(?<mm>\d+)-(?<dd>\d+) (?<hh>\d+):(?<mi>\d+):(?<ss>\d+.\d+)/ =~ datetime_string

          ::Time.new(yyyy, mm, dd, hh, mi, ss.to_r).round(0)
        end
      end
    end
  end
end
