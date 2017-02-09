module Roo
  class Excelx
    module Format
      EXCEPTIONAL_FORMATS = {
        'h:mm am/pm' => :date,
        'h:mm:ss am/pm' => :date
      }

      STANDARD_FORMATS = {
        0 => 'General'.freeze,
        1 => '0'.freeze,
        2 => '0.00'.freeze,
        3 => '#,##0'.freeze,
        4 => '#,##0.00'.freeze,
        9 => '0%'.freeze,
        10 => '0.00%'.freeze,
        11 => '0.00E+00'.freeze,
        12 => '# ?/?'.freeze,
        13 => '# ??/??'.freeze,
        14 => 'mm-dd-yy'.freeze,
        15 => 'd-mmm-yy'.freeze,
        16 => 'd-mmm'.freeze,
        17 => 'mmm-yy'.freeze,
        18 => 'h:mm AM/PM'.freeze,
        19 => 'h:mm:ss AM/PM'.freeze,
        20 => 'h:mm'.freeze,
        21 => 'h:mm:ss'.freeze,
        22 => 'm/d/yy h:mm'.freeze,
        37 => '#,##0 ;(#,##0)'.freeze,
        38 => '#,##0 ;[Red](#,##0)'.freeze,
        39 => '#,##0.00;(#,##0.00)'.freeze,
        40 => '#,##0.00;[Red](#,##0.00)'.freeze,
        45 => 'mm:ss'.freeze,
        46 => '[h]:mm:ss'.freeze,
        47 => 'mmss.0'.freeze,
        48 => '##0.0E+0'.freeze,
        49 => '@'.freeze
      }

      def to_type(format)
        format = format.to_s.downcase
        if (type = EXCEPTIONAL_FORMATS[format])
          type
        elsif format.include?('#')
          :float
        elsif !format.match(/d+(?![\]])/).nil? || format.include?('y')
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

      module_function :to_type
    end
  end 
end
