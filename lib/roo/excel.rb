require 'rubygems'
gem 'parseexcel', '>= 0.5.2'
require 'parseexcel'
CHARGUESS = false
require 'charguess' if CHARGUESS

module Spreadsheet # :nodoc
  module ParseExcel
    class Worksheet
      include Enumerable
      attr_reader :min_row, :max_row, :min_col, :max_col
    end
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
      @workbook = Spreadsheet::ParseExcel.parse(filename)
      @default_sheet = nil
      # no need to set default_sheet if there is only one sheet in the document
      if self.sheets.size == 1
        @default_sheet = self.sheets.first
      end
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
  end

  # returns an array of sheet names in the spreadsheet
  def sheets
    result = []
    #0.upto(@workbook.worksheets.size - 1) do |i| # spreadsheet
    0.upto(@workbook.sheet_count - 1) do |i| # parseexcel
      # TODO: is there a better way to do conversion?
      if CHARGUESS
        encoding = CharGuess::guess(@workbook.worksheet(i).name)
        encoding = 'unicode' unless encoding


        result << Iconv.new('utf-8',encoding).iconv(
          @workbook.worksheet(i).name
        )
      else
        result << platform_specific_iconv(@workbook.worksheet(i).name)
      end
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

  # returns the first non empty column
  def first_column(sheet=nil)
    sheet = @default_sheet unless sheet
    return @first_column[sheet] if @first_column[sheet]
    fr, lr, fc, lc = get_firsts_lasts(sheet)
    fc
  end

  # returns the last non empty column
  def last_column(sheet=nil)
    sheet = @default_sheet unless sheet
    return @last_column[sheet] if @last_column[sheet]
    fr, lr, fc, lc = get_firsts_lasts(sheet)
    lc
  end

  # returns the first non empty row
  def first_row(sheet=nil)
    sheet = @default_sheet unless sheet
    return @first_row[sheet] if @first_row[sheet]
    fr, lr, fc, lc = get_firsts_lasts(sheet)
    fr
  end

  # returns the last non empty row
  def last_row(sheet=nil)
    sheet = @default_sheet unless sheet
    return @last_row[sheet] if @last_row[sheet]
    fr, lr, fc, lc = get_firsts_lasts(sheet)
    lr
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

  # shows the internal representation of all cells
  # mainly for debugging purposes
  def to_s(sheet=nil)
    sheet = @default_sheet unless sheet
    read_cells(sheet) unless @cells_read[sheet]
    @cell[sheet].inspect
  end

  private
  # determine the first and last boundaries
  def get_firsts_lasts(sheet=nil)
    
    # 2008-09-14 BEGINf
    fr=lr=fc=lc=nil
    sheet = @default_sheet unless sheet
    if ! @cells_read[sheet]
      read_cells(sheet)
    end
    if @cell[sheet] # nur wenn ueberhaupt Zellen belegt sind
      @cell[sheet].each {|cellitem|
        key = cellitem.first
        y,x = key

        if cellitem[1].class != String or
            (cellitem[1].class == String and cellitem[1] != "")
          fr = y unless fr
          fr = y if y < fr  
      
          lr = y unless lr
          lr = y if y > lr 
      
          fc = x unless fc
          fc = x if x < fc
      
          lc = x unless lc
          lc = x if x > lc 
        end
      }
    end
    @first_row[sheet]    = fr
    @last_row[sheet]     = lr
    @first_column[sheet] = fc
    @last_column[sheet]  = lc
    return fr, lr, fc, lc
  end

  # converts name of a sheet to index (0,1,2,..)
  def sheet_no(name)
    return name-1 if name.kind_of?(Fixnum)
    0.upto(@workbook.sheet_count - 1) do |i|
      #0.upto(@workbook.worksheets.size - 1) do |i|
      # TODO: is there a better way to do conversion?
      return i if name == platform_specific_iconv(
        @workbook.worksheet(i).name)
      #Iconv.new('utf-8','unicode').iconv(
      #        @workbook.worksheet(i).name
      #      )
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
  def set_cell_values(sheet,x,y,i,v,vt,formula,tr,str_v)
    #key = "#{y},#{x+i}"
    key = [y,x+i]
    @cell_type[sheet] = {} unless @cell_type[sheet]
    @cell_type[sheet][key] = vt 
    @formula[sheet] = {} unless @formula[sheet]
    @formula[sheet][key] = formula  if formula
    @cell[sheet]    = {} unless @cell[sheet]
    case vt # @cell_type[sheet][key]
    when :float
      @cell[sheet][key] = v.to_f
    when :string
      @cell[sheet][key] = str_v
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
    skip = 0
    x =1
    y=1
    i=0
    worksheet.each(skip) { |row_par| 
      if row_par
        x =1
        row_par.each do # |void|
          cell = row_par.at(x-1)
          if cell
            case cell.type
            when :numeric
              vt = :float
              v = cell.to_f
            when :text
              vt = :string
              str_v = cell.to_s('utf-8')
            when :date
              if cell.to_s.to_f < 1.0
                vt = :time
                f = cell.to_s.to_f*24.0*60.0*60.0
                secs = f.round
                h = (secs / 3600.0).floor
                secs = secs - 3600*h
                m = (secs / 60.0).floor
                secs = secs - 60*m
                s = secs
                v = h*3600+m*60+s
              else
                if cell.datetime.hour != 0 or
                    cell.datetime.min  != 0 or
                    cell.datetime.sec  != 0 or
                    cell.datetime.msec != 0
                  vt = :datetime
                  v = cell.datetime
                else
                  vt = :date
                  v = cell.date
                  v = sprintf("%04d-%02d-%02d",v.year,v.month,v.day)
                end
              end
            else
              vt = cell.type.to_s.downcase.to_sym
              v = nil
            end # case
            formula = tr = nil #TODO:???
            set_cell_values(sheet,x,y,i,v,vt,formula,tr,str_v)
          end # if cell
          
          x += 1
        end
      end
      y += 1
    }
    @cells_read[sheet] = true
  end

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
