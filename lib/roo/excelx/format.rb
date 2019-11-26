# frozen_string_literal: true

module Roo
  class Excelx
    module Format
      extend self
      EXCEPTIONAL_FORMATS = {
        'h:mm am/pm' => :date,
        'h:mm:ss am/pm' => :date
      }

      STANDARD_FORMATS = {
        0 => 'General',
        1 => '0',
        2 => '0.00',
        3 => '#,##0',
        4 => '#,##0.00',
        9 => '0%',
        10 => '0.00%',
        11 => '0.00E+00',
        12 => '# ?/?',
        13 => '# ??/??',
        14 => 'mm-dd-yy',
        15 => 'd-mmm-yy',
        16 => 'd-mmm',
        17 => 'mmm-yy',
        18 => 'h:mm AM/PM',
        19 => 'h:mm:ss AM/PM',
        20 => 'h:mm',
        21 => 'h:mm:ss',
        22 => 'm/d/yy h:mm',
        37 => '#,##0 ;(#,##0)',
        38 => '#,##0 ;[Red](#,##0)',
        39 => '#,##0.00;(#,##0.00)',
        40 => '#,##0.00;[Red](#,##0.00)',
        45 => 'mm:ss',
        46 => '[h]:mm:ss',
        47 => 'mmss.0',
        48 => '##0.0E+0',
        49 => '@'
      }

      def to_type(format)
        @to_type ||= {}
        @to_type[format] ||= _to_type(format)
      end

      def _to_type(format)
        format = format.to_s.downcase
        if (type = EXCEPTIONAL_FORMATS[format])
          type
        elsif format.include?('#')
          :float
        elsif format.include?('y') || !format.match(/d+(?![\]])/).nil?
          if format.include?('h') || format.include?('s')
            :datetime
          else
            :date
          end
        elsif format.include?('h') || format.include?('s')
          :time
        elsif format.include?('%')
          :percentage
        else
          :float
        end
      end

    end
  end
end
