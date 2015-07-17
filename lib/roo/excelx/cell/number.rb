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
          case @format
          when /%/
            Float(number)
          when /\.0/
            Float(number)
          else
            number.include?('.') ? Float(number) : Integer(number)
          end
        end

        def formatted_value
          formatter = formats[@format]
          if formatter.is_a? Proc
            formatter.call(@cell_value)
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
            '#,##0' => proc do |number|
              Kernel.format('%.0f', number).reverse.gsub(/(\d{3})(?=\d)/, '\\1,').reverse
            end,
            '#,##0.00' => proc do |number|
              Kernel.format('%.2f', number).reverse.gsub(/(\d{3})(?=\d)/, '\\1,').reverse
            end,
            '0%' =>  proc do |number|
              Kernel.format('%d%', number.to_f * 100)
            end,
            '0.00%' => proc do |number|
              Kernel.format('%.2f%', number.to_f * 100)
            end,
            '0.00E+00' => '%.2E',
            '#,##0 ;(#,##0)' => proc do |number|
              formatter = number.to_i > 0 ? '%.0f' : '(%.0f)'
              Kernel.format(formatter, number.to_f.abs).reverse.gsub(/(\d{3})(?=\d)/, '\\1,').reverse
            end,
            '#,##0 ;[Red](#,##0)' => proc do |number|
              formatter = number.to_i > 0 ? '%.0f' : '[Red](%.0f)'
              Kernel.format(formatter, number.to_f.abs).reverse.gsub(/(\d{3})(?=\d)/, '\\1,').reverse
            end,
            '#,##0.00;(#,##0.00)' => proc do |number|
              formatter = number.to_i > 0 ? '%.2f' : '(%.2f)'
              Kernel.format(formatter, number.to_f.abs).reverse.gsub(/(\d{3})(?=\d)/, '\\1,').reverse
            end,
            '#,##0.00;[Red](#,##0.00)' => proc do |number|
              formatter = number.to_i > 0 ? '%.2f' : '[Red](%.2f)'
              Kernel.format(formatter, number.to_f.abs).reverse.gsub(/(\d{3})(?=\d)/, '\\1,').reverse
            end,
            # FIXME: not quite sure what the format should look like in this case.
            '##0.0E+0' => '%.1E',
            '@' => proc { |number| number }
          }
        end
      end
    end
  end
end
