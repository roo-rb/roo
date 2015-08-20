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
          date_regex = /(?<date>[dmy]+[\-\/][dmy]+([\-\/][dmy]+)?)/
          time_regex = /(?<time>(\[?[h]\]?+:)?[m]+(:?ss|:?s)?)/

          formatter = @format.downcase.split(' ').map do |part|
            if part[date_regex] == part
              part.gsub(/#{DATE_FORMATS.keys.join('|')}/, DATE_FORMATS)
            elsif part[time_regex]
              part.gsub(/#{TIME_FORMATS.keys.join('|')}/, TIME_FORMATS)
            else
              warn 'Unable to parse custom format. Using "YYYY-mm-dd HH:MM:SS" format.'
              return @value.strftime('%F %T')
            end
          end.join(' ')

          @value.strftime(formatter)
        end

        private

        DATE_FORMATS = {
          'yyyy'.freeze => '%Y'.freeze,  # Year: 2000
          'yy'.freeze => '%y'.freeze,    # Year: 00
          # mmmmm => J-D
          'mmmm'.freeze => '%B'.freeze,  # Month: January
          'mmm'.freeze => '%^b'.freeze,   # Month: JAN
          'mm'.freeze => '%m'.freeze,    # Month: 01
          'm'.freeze => '%-m'.freeze,    # Month: 1
          'dddd'.freeze => '%A'.freeze,  # Day of the Week: Sunday
          'ddd'.freeze => '%^a'.freeze,   # Day of the Week: SUN
          'dd'.freeze => '%d'.freeze,    # Day of the Month: 01
          'd'.freeze => '%-d'.freeze,    # Day of the Month: 1
          # '\\\\'.freeze => ''.freeze,  # NOTE: Fixes a custom format's output.
        }

        TIME_FORMATS = {
          'hh'.freeze => '%H'.freeze,    # Hour (24): 01
          'h'.freeze => '%-k'.freeze,    # Hour (24): 1
          # 'hh'.freeze => '%I'.freeze,    # Hour (12): 08
          # 'h'.freeze => '%-l'.freeze,    # Hour (12): 8
          'mm'.freeze => '%M'.freeze,    # Minute: 01
          # FIXME: is this used? Seems like 'm' is used for month, not minute.
          'm'.freeze => '%-M'.freeze,    # Minute: 1
          'ss'.freeze => '%S'.freeze,    # Seconds: 01
          's'.freeze => '%-S'.freeze,    # Seconds: 1
          'am/pm'.freeze => '%p'.freeze, # Meridian: AM
          '000'.freeze => '%3N'.freeze,  # Fractional Seconds: thousandth.
          '00'.freeze => '%2N'.freeze,   # Fractional Seconds: hundredth.
          '0'.freeze => '%1N'.freeze,    # Fractional Seconds: tenths.
        }

        def create_datetime(base_date, value)
          date = base_date + value.to_f.round(6)
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
end
