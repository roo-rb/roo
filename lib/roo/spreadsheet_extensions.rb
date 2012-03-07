require 'spreadsheet'


module Spreadsheet
  # patch for skipping blank rows in the case of
  # having a spreadsheet with 30,000 nil rows appended 
  # to the actual data.  (it happens and your RAM will love me)
  class Worksheet
    def each skip=dimensions[0]
      blanks = 0
      skip.upto(dimensions[1] - 1) do |i|
        if row(i).any?
          Proc.new.call(row(i))
        else
          blanks += 1
          blanks < 20 ? next : return
        end
      end
    end
  end

  module Excel
    class Row < Spreadsheet::Row
      # The Spreadsheet library has a bug in handling Excel 
      # base dates so if the file is a 1904 base date then 
      # dates are off by a day. 1900 base dates work fine
      def _date data # :nodoc:
        return data if data.is_a?(Date)
        date = @worksheet.date_base + data.to_i
        if LEAP_ERROR > @worksheet.date_base
          date -= 1
        end
        date
      end
      public :_datetime

      #=====================================================================
      # TODO:
      # redefinition of this method, the method in the spreadsheet gem has a bug
      # redefinition can be removed, if spreadsheet does it in the correct way
      def _datetime data # :nodoc:
        return data if data.is_a?(DateTime)
        base = @worksheet.date_base
        date = base + data.to_f
        hour = (data % 1) * 24
        min  = (hour % 1) * 60
        sec  = ((min % 1) * 60).round
        min = min.floor
        hour = hour.floor
        if sec > 59
          sec = 0
          min += 1
        end
        if min > 59
          min = 0
          hour += 1
        end
        if hour > 23
          hour = 0
          date += 1
        end
        if LEAP_ERROR > base
          date -= 1
        end
        DateTime.new(date.year, date.month, date.day, hour, min, sec)
      end
    end
  end
end

# ruby-spreadsheet has a font object so we're extending it 
# with our own functionality but still providing full access
# to the user for other font information
module ExcelFontExtensions
  def bold?(*args)
    #From ruby-spreadsheet doc: 100 <= weight <= 1000, bold => 700, normal => 400
    case weight
    when 700    
      true
    else
      false
    end   
  end

  def italic?
    italic
  end

  def underline?
    underline != :none
  end

end
