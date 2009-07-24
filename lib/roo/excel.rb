require 'spreadsheet'
CHARGUESS = begin
  require 'charguess'
  true
rescue LoadError => e
  false
end

# The Spreadsheet library has a bug in handling Excel 
# base dates so if the file is a 1904 base date then 
# dates are off by a day. 1900 base dates work fine
module Spreadsheet
  module Excel
    class Row < Spreadsheet::Row
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
      public :_date
      public :_datetime
    end
    # patch for ruby-spreadsheet parsing formulas
    class Reader
      def read_formula worksheet, addr, work
        row, column, xf, rtype, rval, rcheck, opts = work.unpack 'v3CxCx3v2'
        formula = Formula.new
        formula.shared = (opts & 0x08) > 0
        formula.data = work[20..-1]
        if rcheck != 0xffff || rtype > 3 
          value, = work.unpack 'x6E'
          unless value
            # on architectures where sizeof(double) > 8
            value, = work.unpack 'x6e'
          end
          formula.value = value
        elsif rtype == 0
          pos, op, len, work = get_next_chunk
          if op == :string
            formula.value = client read_string(work, 2), @workbook.encoding
          else
            # This seems to work but I don't know why :). It at least
            # seems to correct the case we saw but doubtful it's the right fix
            formula.value = client read_string(work[10..-1], 2), @workbook.encoding
          end
        elsif rtype == 1
          formula.value = rval > 0
        elsif rtype == 2
          formula.value = Error.new rval
        else
          # leave the Formula value blank
        end
        set_cell worksheet, row, column, xf, formula
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

# Class for handling Excel-Spreadsheets
class Excel < GenericSpreadsheet 

  EXCEL_NO_FORMULAS = 'formulas are not supported for excel spreadsheets'

  # Creates a new Excel spreadsheet object.
  # Parameter packed: :zip - File is a zip-file
  def initialize(filename, packed = nil, file_warning = :error)
    super()
    @file_warning = file_warning
    @tmpdir = "oo_"+$$.to_s
    @tmpdir = File.join(ENV['ROO_TMP'], @tmpdir) if ENV['ROO_TMP'] 
    unless File.exists?(@tmpdir)
      FileUtils::mkdir(@tmpdir)
    end
    filename = open_from_uri(filename) if filename[0,7] == "http://"
    filename = open_from_stream(filename[7..-1]) if filename[0,7] == "stream:"
    filename = unzip(filename) if packed and packed == :zip
    begin
      file_type_check(filename,'.xls','an Excel')
      @filename = filename
      unless File.file?(@filename)
        raise IOError, "file #{@filename} does not exist"
      end
      @workbook = Spreadsheet.open(filename)
      @default_sheet = self.sheets.first
    ensure
      #if ENV["roo_local"] != "thomas-p"
      FileUtils::rm_r(@tmpdir)
      #end
    end
    @cell = Hash.new
    @cell_type = Hash.new
    @formula = Hash.new
    @first_row = Hash.new
    @last_row = Hash.new
    @first_column = Hash.new
    @last_column = Hash.new
    @header_line = 1
    @cells_read = Hash.new
    @fonts = Hash.new
  end

  # returns an array of sheet names in the spreadsheet
  def sheets
    result = []
    @workbook.worksheets.each do |worksheet| 
      result << normalize_string(worksheet.name)
    end
    return result
  end

  # returns the content of a cell. The upper left corner is (1,1) or ('A',1)
  def cell(row,col,sheet=nil)
    sheet = @default_sheet unless sheet
    raise ArgumentError unless sheet
    read_cells(sheet) unless @cells_read[sheet]
    raise "should be read" unless @cells_read[sheet]
    row,col = normalize(row,col)
    if celltype(row,col,sheet) == :date
      yyyy,mm,dd = @cell[sheet][[row,col]].split('-')
      return Date.new(yyyy.to_i,mm.to_i,dd.to_i)
    end
    if celltype(row,col,sheet) == :string
      return platform_specific_iconv(@cell[sheet][[row,col]])
    else
      return @cell[sheet][[row,col]]
    end
  end

  # returns the type of a cell:
  # * :float
  # * :string,
  # * :date
  # * :percentage
  # * :formula
  # * :time
  # * :datetime
  def celltype(row,col,sheet=nil)
    sheet = @default_sheet unless sheet
    read_cells(sheet) unless @cells_read[sheet]
    row,col = normalize(row,col)
    begin
      if @formula[sheet][[row,col]]
        return :formula
      else
        @cell_type[sheet][[row,col]]
      end
    rescue
      puts "Error in sheet #{sheet}, row #{row}, col #{col}"
      raise
    end
  end

  # returns NO formula in excel spreadsheets
  def formula(row,col,sheet=nil)
    raise EXCEL_NO_FORMULAS
  end

  # raises an exception because formulas are not supported for excel files
  def formula?(row,col,sheet=nil)
    raise EXCEL_NO_FORMULAS
  end

  # returns NO formulas in excel spreadsheets
  def formulas(sheet=nil)
    raise EXCEL_NO_FORMULAS
  end

  # Given a cell, return the cell's font
  def font(row, col, sheet=nil)
    sheet = @default_sheet unless sheet
    read_cells(sheet) unless @cells_read[sheet]
    row,col = normalize(row,col)
    @fonts[sheet][[row,col]]
  end 
  
  # shows the internal representation of all cells
  # mainly for debugging purposes
  def to_s(sheet=nil)
    sheet = @default_sheet unless sheet
    read_cells(sheet) unless @cells_read[sheet]
    @cell[sheet].inspect
  end

  private

  # converts name of a sheet to index (0,1,2,..)
  def sheet_no(name)
    return name-1 if name.kind_of?(Fixnum)
    i = 0
    @workbook.worksheets.each do |worksheet|
      return i if name == normalize_string(worksheet.name)
      i += 1
    end
    raise StandardError, "sheet '#{name}' not found"
  end

  def empty_row?(row)
    content = false
    row.compact.each {|elem|
      if elem != ''
        content = true
      end
    }
    ! content
  end

  def empty_column?(col)
    content = false
    col.compact.each {|elem|
      if elem != ''
        content = true
      end
    }
    ! content
  end
  
  def normalize_string(value)
    value = every_second_null?(value) ? remove_every_second_null(value) : value
    if CHARGUESS && encoding = CharGuess::guess(value)
      Iconv.new('utf-8', encoding)
    else
      platform_specific_iconv(value)
    end
  end
  
  def platform_specific_iconv(value)
    case RUBY_PLATFORM.downcase
    when /darwin/
      result = Iconv.new('utf-8','utf-8').iconv(value)
    when /solaris/
      result = Iconv.new('utf-8','utf-8').iconv(value)
    when /mswin32/
      result = Iconv.new('utf-8','iso-8859-1').iconv(value)
    else
      result = value
    end # case
    if every_second_null?(result)
      result = remove_every_second_null(result)
    end
    result
  end

  def every_second_null?(str)
    result = true
    return false if str.length < 2
    0.upto(str.length/2-1) do |i|
      c = str[i*2,1]
      n = str[i*2+1,1]
      if n != "\000"
        result = false 
        break
      end
    end
    result
  end

  def remove_every_second_null(str)
    result = ''
    0.upto(str.length/2-1) do |i|
      c = str[i*2,1]
      result += c
    end
    result
  end

  # helper function to set the internal representation of cells
  def set_cell_values(sheet,row,col,i,v,vt,formula,tr,font)
    #key = "#{y},#{x+i}"
    key = [row,col+i]
    @cell_type[sheet] = {} unless @cell_type[sheet]
    @cell_type[sheet][key] = vt 
    @formula[sheet] = {} unless @formula[sheet]
    @formula[sheet][key] = formula  if formula
    @cell[sheet]    = {} unless @cell[sheet]
    @fonts[sheet] = {} unless @fonts[sheet]
    @fonts[sheet][key] = font
    
    case vt # @cell_type[sheet][key]
    when :float
      @cell[sheet][key] = v.to_f
    when :string
      @cell[sheet][key] = v
    when :date
      @cell[sheet][key] = v 
    when :datetime
      @cell[sheet][key] = DateTime.new(v.year,v.month,v.day,v.hour,v.min,v.sec)
    when :percentage
      @cell[sheet][key] = v.to_f
    when :time
      @cell[sheet][key] = v 
    else
      @cell[sheet][key] = v
    end
  end

  # read all cells in the selected sheet
  def read_cells(sheet=nil)
    sheet = @default_sheet unless sheet
    raise ArgumentError, "Error: sheet '#{sheet||'nil'}' not valid" if @default_sheet == nil and sheet==nil
    raise RangeError unless self.sheets.include? sheet
    
    if @cells_read[sheet]
      raise "sheet #{sheet} already read"
    end
    
    worksheet = @workbook.worksheet(sheet_no(sheet))
    row_index=1
    worksheet.each(0) do |row| 
      (0..row.size).each do |cell_index|
        cell = row.at(cell_index)
        next if cell.nil?  #skip empty cells
        next if cell.class == Spreadsheet::Formula && cell.value.nil? # skip empty formla cells
        if date_or_time?(row, cell_index)
          vt, v = read_cell_date_or_time(row, cell_index)
        else
          vt, v = read_cell(row, cell_index)
        end
        formula = tr = nil #TODO:???
        col_index = cell_index + 1
        font = row.format(cell_index).font
        font.extend(ExcelFontExtensions)
        set_cell_values(sheet,row_index,col_index,0,v,vt,formula,tr,font)
      end #row
      row_index += 1
    end # worksheet
    @cells_read[sheet] = true
  end

  # Get the contents of a cell, accounting for the
  # way formula stores the value
  def read_cell_content(row, idx)
    cell = row.at(idx)
    cell = cell.value if cell.class == Spreadsheet::Formula
    cell
  end
    
  # Test the cell to see if it's a valid date/time. 
  def date_or_time?(row, idx)
    format = row.format(idx)
    if format.date_or_time?
      cell = read_cell_content(row, idx)
      true if Float(cell) > 0 rescue false
    else
      false
    end  
  end
  private :date_or_time?
  
  # Read the date-time cell and convert to, 
  # the date-time values for Roo
  def read_cell_date_or_time(row, idx)
    cell = read_cell_content(row, idx)
    cell = cell.to_s.to_f
    if cell < 1.0
      value_type = :time
      f = cell*24.0*60.0*60.0
      secs = f.round
      h = (secs / 3600.0).floor
      secs = secs - 3600*h
      m = (secs / 60.0).floor
      secs = secs - 60*m
      s = secs
      value = h*3600+m*60+s
    else
      if row.at(idx).class == Spreadsheet::Formula
        datetime = row._datetime(cell)
      else
        datetime = row.datetime(idx)
      end    
      if datetime.hour != 0 or
         datetime.min != 0 or
         datetime.sec != 0 
        value_type = :datetime
        value = datetime
      else
        value_type = :date
        if row.at(idx).class == Spreadsheet::Formula
          value = row._date(cell)
        else
          value = row.date(idx)
        end    
        value = sprintf("%04d-%02d-%02d",value.year,value.month,value.day)
      end
    end  
    return value_type, value
  end
  private :read_cell_date_or_time
  
  # Read the cell and based on the class, 
  # return the values for Roo
  def read_cell(row, idx)
    cell = read_cell_content(row, idx)
    case cell
    when Float, Integer, Fixnum, Bignum 
      value_type = :float
      value = cell.to_f
    when String, TrueClass, FalseClass
      value_type = :string
      value = cell.to_s
    else
      value_type = cell.class.to_s.downcase.to_sym
      value = nil
    end # case
    return value_type, value
  end
  private :read_cell
  
  #TODO: testing only
  #  def inject_null_characters(str)
  #    if str.class != String
  #      return str
  #    end
  #    new_str=''
  #    0.upto(str.size-1) do |i|
  #      new_str += str[i,1]
  #      new_str += "\000"
  #    end
  #    new_str
  #  end
  #

end
