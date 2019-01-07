# frozen_string_literal: true

module Roo
  class Excelx
    class Cell
      class Number < Cell::Base
        attr_reader :value, :formula, :format, :cell_value, :coordinate

        # FIXME: change default_type to number. This will break brittle tests.
        attr_reader_with_default default_type: :float

        def initialize(value, formula, excelx_type, style, link, coordinate)
          super
          # FIXME: Excelx_type is an array, but the first value isn't used.
          @format = excelx_type.last
          @value = link ? Roo::Link.new(link, value) : create_numeric(value)
        end

        def create_numeric(number)
          return number if Excelx::ERROR_VALUES.include?(number)
          case @format
          when /%/
            Float(number)
          when /\.0/
            Float(number)
          else
            (number.include?('.') || (/\A[-+]?\d+E[-+]?\d+\z/i =~ number)) ? Float(number) : Integer(number, 10)
          end
        end

        def formatted_value
          return @cell_value if Excelx::ERROR_VALUES.include?(@cell_value)

          formatter = generate_formatter(@format)
          if formatter.is_a? Proc
            formatter.call(@cell_value)
          else
            Kernel.format(formatter, @cell_value)
          end
        end

        def generate_formatter(format)
          # FIXME: numbers can be other colors besides red:
          # [BLACK], [BLUE], [CYAN], [GREEN], [MAGENTA], [RED], [WHITE], [YELLOW], [COLOR n]
          case format
          when /^General$/i then '%.0f'
          when '0' then '%.0f'
          when /^(0+)$/ then "%0#{$1.size}d"
          when /^0\.(0+)$/ then "%.#{$1.size}f"
          when '#,##0' then number_format('%.0f')
          when '#,##0.00' then number_format('%.2f')
          when '0%'
            proc do |number|
              Kernel.format('%d%%', number.to_f * 100)
            end
          when '0.00%'
            proc do |number|
              Kernel.format('%.2f%%', number.to_f * 100)
            end
          when '0.00E+00' then '%.2E'
          when '#,##0 ;(#,##0)' then number_format('%.0f', '(%.0f)')
          when '#,##0 ;[Red](#,##0)' then number_format('%.0f', '[Red](%.0f)')
          when '#,##0.00;(#,##0.00)' then number_format('%.2f', '(%.2f)')
          when '#,##0.00;[Red](#,##0.00)' then number_format('%.2f', '[Red](%.2f)')
            # FIXME: not quite sure what the format should look like in this case.
          when '##0.0E+0' then '%.1E'
          when '@' then proc { |number| number }
          else
            raise "Unknown format: #{format.inspect}"
          end
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
      end
    end
  end
end
