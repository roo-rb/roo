#require 'xml'
require "google_spreadsheet"

class GoogleHTTPError < RuntimeError; end
class GoogleReadError < RuntimeError; end
class GoogleWriteError < RuntimeError; end

class Google < GenericSpreadsheet
  attr_accessor :date_format, :datetime_format
  
  # Creates a new Google spreadsheet object.
  def initialize(spreadsheetkey,user=nil,password=nil)
    @filename = spreadsheetkey
    @spreadsheetkey = spreadsheetkey
    @user = user
    @password = password
    unless user
      user = ENV['GOOGLE_MAIL']
    end
    unless password
      password = ENV['GOOGLE_PASSWORD']
    end
    @cell = Hash.new {|h,k| h[k]=Hash.new}
    @cell_type = Hash.new {|h,k| h[k]=Hash.new}
    @formula = Hash.new
    @first_row = Hash.new
    @last_row = Hash.new
    @first_column = Hash.new
    @last_column = Hash.new
    @cells_read = Hash.new
    @header_line = 1
    @date_format = '%d/%m/%Y'
    @datetime_format = '%d/%m/%Y %H:%M:%S' 
    @time_format = '%H:%M:%S'
    session = GoogleSpreadsheet.login(user, password)
    @sheetlist = []
    session.spreadsheet_by_key(@spreadsheetkey).worksheets.each { |sheet|
      @sheetlist << sheet.title
    }
    @default_sheet = self.sheets.first
    @worksheets = session.spreadsheet_by_key(@spreadsheetkey).worksheets
  end

  # returns an array of sheet names in the spreadsheet
  def sheets
    @sheetlist
  end

  def date?(string)
    begin
      Date.strptime(string, @date_format)
      true
    rescue
      false
    end
  end

  # is String a time with format HH:MM:SS?
  def time?(string)
    begin
      DateTime.strptime(string, @time_format)
      true
    rescue
      false
    end
  end

  def datetime?(string)
    begin
      DateTime.strptime(string, @datetime_format)
      true
    rescue
      false
    end
  end

  def numeric?(string)
    string =~ /^[0-9]+[\.]*[0-9]*$/
  end

  def timestring_to_seconds(value)
    hms = value.split(':')
    hms[0].to_i*3600 + hms[1].to_i*60 + hms[2].to_i
  end

  # Returns the content of a spreadsheet-cell.
  # (1,1) is the upper left corner.
  # (1,1), (1,'A'), ('A',1), ('a',1) all refers to the
  # cell at the first line and first row.
  def cell(row, col, sheet=nil)
    sheet = @default_sheet unless sheet
    check_default_sheet #TODO: 2007-12-16
    read_cells(sheet) unless @cells_read[sheet]
    row,col = normalize(row,col)
    value = @cell[sheet]["#{row},#{col}"]
    if celltype(row,col,sheet) == :date
      begin
        return  Date.strptime(value, @date_format)
      rescue ArgumentError
        raise "Invalid Date #{sheet}[#{row},#{col}] #{value} using format '{@date_format}'"
      end
    elsif celltype(row,col,sheet) == :datetime
      begin
        return  DateTime.strptime(value, @datetime_format)
      rescue ArgumentError
        raise "Invalid DateTime #{sheet}[#{row},#{col}] #{value} using format '{@datetime_format}'"
      end
    end 
    return value
  end

  # returns the type of a cell:
  # * :float
  # * :string
  # * :date
  # * :percentage
  # * :formula
  # * :time
  # * :datetime
  def celltype(row, col, sheet=nil)
    sheet = @default_sheet unless sheet
    read_cells(sheet) unless @cells_read[sheet]
    row,col = normalize(row,col)
    if @formula.size > 0 && @formula[sheet]["#{row},#{col}"]
      return :formula
    else
      @cell_type[sheet]["#{row},#{col}"]
    end
  end

  # Returns the formula at (row,col).
  # Returns nil if there is no formula.
  # The method #formula? checks if there is a formula.
  def formula(row,col,sheet=nil)
    sheet = @default_sheet unless sheet
    read_cells(sheet) unless @cells_read[sheet]
    row,col = normalize(row,col)
    if @formula[sheet]["#{row},#{col}"] == nil
      return nil
    else
      return @formula[sheet]["#{row},#{col}"] 
    end
  end

  # true, if there is a formula
  def formula?(row,col,sheet=nil)
    sheet = @default_sheet unless sheet
    read_cells(sheet) unless @cells_read[sheet]
    row,col = normalize(row,col)
    formula(row,col) != nil
  end

  # returns each formula in the selected sheet as an array of elements
  # [row, col, formula]
  def formulas(sheet=nil)
    theformulas = Array.new
    sheet = @default_sheet unless sheet
    read_cells(sheet) unless @cells_read[sheet]
    first_row(sheet).upto(last_row(sheet)) {|row|
      first_column(sheet).upto(last_column(sheet)) {|col|
        if formula?(row,col,sheet)
          f = [row, col, formula(row,col,sheet)]
          theformulas << f
        end
      }
    }
    theformulas
  end

  # true, if the cell is empty
  def empty?(row, col, sheet=nil)
    value = cell(row, col, sheet)
    return true unless value
    return false if value.class == Date # a date is never empty
    return false if value.class == Float
    return false if celltype(row,col,sheet) == :time
    value.empty?
  end

  # sets the cell to the content of 'value'
  # a formula can be set in the form of '=SUM(...)'
  def set_value(row,col,value,sheet=nil)
    sheet = @default_sheet unless sheet
    raise RangeError, "sheet not set" unless sheet
    #@@ Set and pass sheet_no
    begin
      sheet_no = sheets.index(sheet)+1
    rescue
      raise RangeError, "invalid sheet '"+sheet.to_s+"'"
    end
    row,col = normalize(row,col)
    add_to_cell_roo(row,col,value,sheet_no)
    # re-read the portion of the document that has changed
    if @cells_read[sheet]
      key = "#{row},#{col}"
      (value, value_type) = determine_datatype(value.to_s)
      @cell[sheet][key] = value 
      @cell_type[sheet][key] = value_type 
    end
  end
  
  # returns the first non-empty row in a sheet
  def first_row(sheet=nil)
    sheet = @default_sheet unless sheet
    unless @first_row[sheet]
      sheet_no = sheets.index(sheet) + 1
      @first_row[sheet], @last_row[sheet], @first_column[sheet], @last_column[sheet] =
          oben_unten_links_rechts(sheet_no)
    end   
    return @first_row[sheet]
  end

  # returns the last non-empty row in a sheet
  def last_row(sheet=nil)
    sheet = @default_sheet unless sheet
    unless @last_row[sheet]
      sheet_no = sheets.index(sheet) + 1
      @first_row[sheet], @last_row[sheet], @first_column[sheet], @last_column[sheet] =
          oben_unten_links_rechts(sheet_no)
    end
    return @last_row[sheet]
  end

  # returns the first non-empty column in a sheet
  def first_column(sheet=nil)
    sheet = @default_sheet unless sheet
    unless @first_column[sheet]
      sheet_no = sheets.index(sheet) + 1
      @first_row[sheet], @last_row[sheet], @first_column[sheet], @last_column[sheet] =
          oben_unten_links_rechts(sheet_no)
    end
    return @first_column[sheet]
  end

  # returns the last non-empty column in a sheet
  def last_column(sheet=nil)
    sheet = @default_sheet unless sheet
    unless @last_column[sheet]
      sheet_no = sheets.index(sheet) + 1
      @first_row[sheet], @last_row[sheet], @first_column[sheet], @last_column[sheet] =
        oben_unten_links_rechts(sheet_no)
    end
    return @last_column[sheet]
  end

  private

  # read all cells in a sheet. 
  def read_cells(sheet=nil)
    sheet = @default_sheet unless sheet
    raise RangeError, "illegal sheet <#{sheet}>" unless sheets.index(sheet)
    sheet_no = sheets.index(sheet)
    ws = @worksheets[sheet_no]
    for row in 1..ws.num_rows
      for col in 1..ws.num_cols
        key = "#{row},#{col}"
        string_value = ws.input_value(row,col) # item['inputvalue'] ||  item['inputValue']
        numeric_value = ws[row,col] #item['numericvalue']  ||  item['numericValue']
        (value, value_type) = determine_datatype(string_value, numeric_value)
        @cell[sheet][key] = value unless value == "" or value == nil
        @cell_type[sheet][key] = value_type
        @formula[sheet] = {} unless @formula[sheet]
        @formula[sheet][key] = string_value if value_type == :formula
        ############
        #$log.debug("key: #{key}")
        #$log.debug "#{ws[row,col].inspect}"
        #@cell[sheet][key] = ws[row,col]
        #$log.debug "@cell[sheet][key]: #{@cell[sheet][key]}"
        ############
      end
    end
    @cells_read[sheet] = true
  end
  
  def determine_datatype(val, numval=nil)
#    $log.debug "val: #{val} numval: #{numval}"
    if val.nil? || val[0,1] == '='
      ty = :formula
      if numeric?(numval)
        val = numval.to_f
      else
        val = numval
      end
    else
      if datetime?(val)
        ty = :datetime
      elsif date?(val)
        ty = :date
      elsif numeric?(val) 
        ty = :float
        val = val.to_f
      elsif time?(val)
        ty = :time
        val = timestring_to_seconds(val)
      else
        ty = :string
      end
    end
    #$log.debug "val: #{val} ty: #{ty}" if ty == :date
    return val, ty 
  end

  def add_to_cell_roo(row,col,value, sheet_no=1)
    sheet_no -= 1
    @worksheets[sheet_no][row,col] = value
    @worksheets[sheet_no].save
  end
  def entry_roo(value,row,col)
    return value,row,col
  end

  def oben_unten_links_rechts(sheet_no)
    ws = @worksheets[sheet_no-1]
    rows = []
    cols = []
    for row in 1..ws.num_rows
      for col in 1..ws.num_cols
        rows << row if ws[row,col] and ws[row,col] != '' #TODO: besser?
        cols << col if ws[row,col] and ws[row,col] != '' #TODO: besser?
      end
    end
    return rows.min, rows.max, cols.min, cols.max
  end

end # class