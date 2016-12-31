module Roo
  class Excelx
    class Cell
      class Number < Cell::Base
        attr_reader :value, :formula, :format, :cell_value, :link, :coordinate

        def initialize(value, formula, excelx_type, style, link, coordinate)
          super
          # FIXME: change @type to number. This will break brittle tests.
          # FIXME: Excelx_type is an array, but the first value isn't used.
          @type = :float
          @format = excelx_type.last
          @value = link? ? Roo::Link.new(link, value) : create_numeric(value)
        end

        def create_numeric(number)
          return number if Excelx::ERROR_VALUES.include?(number)
          case @format
          when /%/
            Float(number)
          when /\.0/
            Float(number)
          else
            (number.include?('.') || (/\A[-+]?\d+E[-+]\d+\z/i =~ number)) ? Float(number) : Integer(number)
          end
        end

        def formatted_value
          return @cell_value if Excelx::ERROR_VALUES.include?(@cell_value)

          formatter = formats[@format]
          if formatter.is_a? Proc
            formatter.call(@cell_value)
          elsif zero_padded_number?
            "%0#{@format.size}d" % @cell_value
          else
            Kernel.format(formatter, @cell_value)
          end
        end

        def formats
          # FIXME: numbers can be other colors besides red:
          # [BLACK], [BLUE], [CYAN], [GREEN], [MAGENTA], [RED], [WHITE], [YELLOW], [COLOR n]
          {
            'General' => '%.0f',
            '0' => '%.0f',
            '0.00' => '%.2f',
            '0.000000' => '%.6f',
            '#,##0' => number_format('%.0f'),
            '#,##0.00' => number_format('%.2f'),
            '0%' =>  proc do |number|
              Kernel.format('%d%', number.to_f * 100)
            end,
            '0.00%' => proc do |number|
              Kernel.format('%.2f%', number.to_f * 100)
            end,
            '0.00E+00' => '%.2E',
            '#,##0 ;(#,##0)' => number_format('%.0f', '(%.0f)'),
            '#,##0 ;[Red](#,##0)' => number_format('%.0f', '[Red](%.0f)'),
            '#,##0.00;(#,##0.00)' => number_format('%.2f', '(%.2f)'),
            '#,##0.00;[Red](#,##0.00)' => number_format('%.2f', '[Red](%.2f)'),
            # FIXME: not quite sure what the format should look like in this case.
            '##0.0E+0' => '%.1E',
            '@' => proc { |number| number }
          }
        end

        private

        def number_format(formatter, negative_formatter = nil)
          proc do |number|
            if negative_formatter
              formatter = number.to_i > 0 ? formatter : negative_formatter
              number = number.to_f.abs
            end

            Kernel.format(formatter, number).reverse.gsub(/(\d{3})(?=\d)/, '\\1,').reverse
          end
        end

        def zero_padded_number?
          @format[/0+/] == @format
        end
      end
    end
  end
end
