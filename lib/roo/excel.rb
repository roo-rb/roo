require 'spreadsheet'

# Class for handling Excel-Spreadsheets
class Roo::Excel < Roo::Base
  FORMULAS_MESSAGE = 'the spreadsheet gem does not support forumulas, so roo can not.'
  CHARGUESS =
    begin
      require 'charguess'
      true
    rescue LoadError
      false
    end

  attr_reader :workbook

  # Creates a new Excel spreadsheet object.
  # Parameter packed: :zip - File is a zip-file
  def initialize(filename, options = {}, deprecated_file_warning = :error)
    if Hash === options
      packed = options[:packed]
      file_warning = options[:file_warning] || :error
      mode = options[:mode] || "rb+"
    else
      warn 'Supplying `packed` or `file_warning` as separate arguments to `Roo::Excel.new` is deprecated. Use an options hash instead.'
      packed = options
      mode = "rb+"
      file_warning = deprecated_file_warning
    end

    file_type_check(filename,'.xls','an Excel', file_warning, packed)
    make_tmpdir do |tmpdir|
      filename = download_uri(filename, tmpdir) if uri?(filename)
      filename = open_from_stream(filename[7..-1], tmpdir) if filename[0,7] == "stream:"
      filename = unzip(filename, tmpdir) if packed == :zip

      @filename = filename
      unless File.file?(@filename)
        raise IOError, "file #{@filename} does not exist"
      end
      @workbook = Spreadsheet.open(filename, mode)
    end
    super(filename, options)
    @formula = Hash.new
    @fonts = Hash.new
  end

  def encoding=(codepage)
    @workbook.encoding = codepage
  end

  # returns an array of sheet names in the spreadsheet
  def sheets
    @workbook.worksheets.collect {|worksheet| normalize_string(worksheet.name)}
  end

  # this method lets you find the worksheet with the most data
  def longest_sheet
    sheet(@workbook.worksheets.inject {|m,o|
      o.row_count > m.row_count ? o : m
    }.name)
  end

  # returns the content of a cell. The upper left corner is (1,1) or ('A',1)
  def cell(row,col,sheet=nil)
    sheet ||= @default_sheet
    validate_sheet!(sheet)

    read_cells(sheet)
    raise "should be read" unless @cells_read[sheet]
    row,col = normalize(row,col)
    if celltype(row,col,sheet) == :date
      yyyy,mm,dd = @cell[sheet][[row,col]].split('-')
      return Date.new(yyyy.to_i,mm.to_i,dd.to_i)
    end
    if celltype(row,col,sheet) == :string
      return platform_specific_encoding(@cell[sheet][[row,col]])
    else
      if @cell[sheet] and @cell[sheet][[row,col]]
        return @cell[sheet][[row,col]]
      else
        return nil
      end
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
    sheet ||= @default_sheet
    read_cells(sheet)
    row,col = normalize(row,col)
    begin
      if @formula[sheet] and @formula[sheet][[row,col]]
        :formula
      elsif @cell_type[sheet]
        @cell_type[sheet][[row,col]]
      end
    rescue
      puts "Error in sheet #{sheet}, row #{row}, col #{col}"
      raise
    end
  end

  # returns NO formula in excel spreadsheets
  def formula(row,col,sheet=nil)
    raise NotImplementedError, FORMULAS_MESSAGE
  end
  alias_method :formula?, :formula

  # returns NO formulas in excel spreadsheets
  def formulas(sheet=nil)
    raise NotImplementedError, FORMULAS_MESSAGE
  end

  # Given a cell, return the cell's font
  def font(row, col, sheet=nil)
    sheet ||= @default_sheet
    read_cells(sheet)
    row,col = normalize(row,col)
    @fonts[sheet][[row,col]]
  end

  # shows the internal representation of all cells
  # mainly for debugging purposes
  def to_s(sheet=nil)
    sheet ||= @default_sheet
    read_cells(sheet)
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

  def normalize_string(value)
    value = every_second_null?(value) ? remove_every_second_null(value) : value
    if CHARGUESS && encoding = CharGuess::guess(value)
      encoding.encode Encoding::UTF_8
    else
      platform_specific_encoding(value)
    end
  end

  def platform_specific_encoding(value)
    result =
      case RUBY_PLATFORM.downcase
      when /darwin|solaris/
        value.encode Encoding::UTF_8
      when /mswin32/
        value.encode Encoding::ISO_8859_1
      else
        value
      end
    if every_second_null?(result)
      result = remove_every_second_null(result)
    end
    result
  end

  def every_second_null?(str)
    result = true
    return false if str.length < 2
    0.upto(str.length/2-1) do |i|
      if str[i*2+1,1] != "\000"
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
  def set_cell_values(sheet,row,col,i,v,value_type,formula,tr,font)
    #key = "#{y},#{x+i}"
    key = [row,col+i]
    @cell_type[sheet] = {} unless @cell_type[sheet]
    @cell_type[sheet][key] = value_type
    @formula[sheet] = {} unless @formula[sheet]
    @formula[sheet][key] = formula  if formula
    @cell[sheet]    = {} unless @cell[sheet]
    @fonts[sheet] = {} unless @fonts[sheet]
    @fonts[sheet][key] = font

    @cell[sheet][key] =
      case value_type
      when :float
        v.to_f
      when :string
        v
      when :date
        v
      when :datetime
        @cell[sheet][key] = DateTime.new(v.year,v.month,v.day,v.hour,v.min,v.sec)
      when :percentage
        v.to_f
      when :time
        v
      else
        v
      end
  end

  # ruby-spreadsheet has a font object so we're extending it
  # with our own functionality but still providing full access
  # to the user for other font information
  module ExcelFontExtensions
    def bold?(*args)
      #From ruby-spreadsheet doc: 100 <= weight <= 1000, bold => 700, normal => 400
      weight == 700
    end

    def italic?
      italic
    end

    def underline?
      underline != :none
    end
  end

  # read all cells in the selected sheet
  def read_cells(sheet=nil)
    sheet ||= @default_sheet
    validate_sheet!(sheet)
    return if @cells_read[sheet]

    worksheet = @workbook.worksheet(sheet_no(sheet))
    row_index=1
    worksheet.each(0) do |row|
      (0..row.size).each do |cell_index|
        cell = row.at(cell_index)
        next if cell.nil?  #skip empty cells
        next if cell.class == Spreadsheet::Formula && cell.value.nil? # skip empty formula cells
        value_type, v =
          if date_or_time?(row, cell_index)
            read_cell_date_or_time(row, cell_index)
          else
            read_cell(row, cell_index)
          end
        formula = tr = nil #TODO:???
        col_index = cell_index + 1
        font = row.format(cell_index).font
        font.extend(ExcelFontExtensions)
        set_cell_values(sheet,row_index,col_index,0,v,value_type,formula,tr,font)
      end #row
      row_index += 1
    end # worksheet
    @cells_read[sheet] = true
  end

  # Get the contents of a cell, accounting for the
  # way formula stores the value
  def read_cell_content(row, idx)
    cell = row.at(idx)
    cell = row[idx] if row[idx].class == Spreadsheet::Link
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
        datetime = row.send(:_datetime, cell)
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
          value = row.send(:_date, cell)
        else
          value = row.date(idx)
        end
        value = sprintf("%04d-%02d-%02d",value.year,value.month,value.day)
      end
    end
    return value_type, value
  end

  # Read the cell and based on the class,
  # return the values for Roo
  def read_cell(row, idx)
    cell = read_cell_content(row, idx)
    case cell
    when Float, Integer, Fixnum, Bignum
      value_type = :float
      value = cell.to_f
    when Spreadsheet::Link
      value_type = :link
      value = cell
    when String, TrueClass, FalseClass
      value_type = :string
      value = cell.to_s
    else
      value_type = cell.class.to_s.downcase.to_sym
      value = nil
    end # case
    return value_type, value
  end

end
