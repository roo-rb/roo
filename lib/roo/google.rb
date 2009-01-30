require 'gdata/spreadsheet'

# overwrite some methods from the gdata-gem:
module GData
  class Spreadsheet < GData::Base
    #-- modified
    def evaluate_cell(cell, sheet_no=1)
      raise ArgumentError, "invalid cell: #{cell}" unless cell
      raise ArgumentError, "invalid sheet_no: #{sheet_no}" unless sheet_no >0 and sheet_no.class == Fixnum
      path = "/feeds/cells/#{@spreadsheet_id}/#{sheet_no}/#{@headers ? "private" : "public"}/basic/#{cell}"

      doc = Hpricot(request(path))
      result = (doc/"content").inner_html
    end

    #-- new
    def sheetlist
      path = "/feeds/worksheets/#{@spreadsheet_id}/private/basic"
      doc = Hpricot(request(path))
      result = []
      (doc/"content").each { |elem|
        result << elem.inner_html
      }
      result
    end

    #-- new
    #@@ added sheet_no to definition
    def save_entry_roo(entry, sheet_no)
      path = "/feeds/cells/#{@spreadsheet_id}/#{sheet_no}/#{@headers ? 'private' : 'public'}/full"
      post(path, entry)
    end

    #-- new
    def entry_roo(formula, row=1, col=1)
      <<XML
    <entry xmlns='http://www.w3.org/2005/Atom' xmlns:gs='http://schemas.google.com/spreadsheets/2006'>
      <gs:cell row='#{row}' col='#{col}' inputValue='#{formula}' />
    </entry>
XML
    end

    #-- new
    #@@ added sheet_no to definition		
    def add_to_cell_roo(row,col,value, sheet_no=1)
      save_entry_roo(entry_roo(value,row,col), sheet_no)
    end

    #-- new
    def get_one_sheet
      path = "/feeds/cells/#{@spreadsheet_id}/1/private/full"
      doc = Hpricot(request(path))
    end

    #new
    def oben_unten_links_rechts(sheet_no)
      path = "/feeds/cells/#{@spreadsheet_id}/#{sheet_no}/private/full"
      doc = Hpricot(request(path))
      rows = []
      cols = []
      (doc/"gs:cell").each {|item|
        rows.push item['row'].to_i
        cols.push item['col'].to_i
      }
      return rows.min, rows.max, cols.min, cols.max
    end

    def fulldoc(sheet_no)
      path = "/feeds/cells/#{@spreadsheet_id}/#{sheet_no}/private/full"
      doc = Hpricot(request(path))
      return doc
    end
  end # class
end # module

class Google < GenericSpreadsheet

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
    @default_sheet = nil
    @cell = Hash.new
    @cell_type = Hash.new
    @formula = Hash.new
    @first_row = Hash.new
    @last_row = Hash.new
    @first_column = Hash.new
    @last_column = Hash.new
    @cells_read = Hash.new
    @header_line = 1

    @gs = GData::Spreadsheet.new(spreadsheetkey)
    @gs.authenticate(user, password)

    #-- ----------------------------------------------------------------------
    #-- TODO: Behandlung von Berechtigungen hier noch einbauen ???
    #-- ----------------------------------------------------------------------

    if self.sheets.size  == 1
      @default_sheet = self.sheets.first
    end
  end

  # returns an array of sheet names in the spreadsheet
  def sheets
    return @gs.sheetlist
  end

  # is String a date with format DD/MM/YYYY
  def Google.date?(string)
    return false if string.class == Float
    return true if string.class == Date
    return string.strip =~ /^([0-9]+)\/([0-9]+)\/([0-9]+)$/
  end

  # is String a time with format HH:MM:SS?
  def Google.time?(string)
    return false if string.class == Float
    return true if string.class == Date
    return string.strip =~ /^([0-9]+):([0-9]+):([0-9]+)$/
  end

  # is String a date+time with format DD/MM/YYYY HH:MM:SS
  def Google.datetime?(string)
    return false if string.class == Float
    return true if string.class == Date
    return string.strip =~ /^([0-9]+)\/([0-9]+)\/([0-9]+)\ ([0-9]+):([0-9]+):([0-9]+)$/
  end


  def Google.timestring_to_seconds(value)
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
    if celltype(row,col,sheet) == :date
      yyyy,mm,dd = @cell[sheet]["#{row},#{col}"].split('-')
      begin
        return Date.new(yyyy.to_i,mm.to_i,dd.to_i)
      rescue ArgumentError
        raise "Invalid date parameter: #{yyyy}, #{mm}, #{dd}"
      end
    elsif celltype(row,col,sheet) == :datetime
      begin
        date_part,time_part = @cell[sheet]["#{row},#{col}"].split(' ')
        yyyy,mm,dd = date_part.split('-')
        hh,mi,ss = time_part.split(':')
        return DateTime.civil(yyyy.to_i,mm.to_i,dd.to_i,hh.to_i,mi.to_i,ss.to_i)
      rescue ArgumentError
        raise "Invalid date parameter: #{yyyy}, #{mm}, #{dd}, #{hh}, #{mi}, #{ss}"
      end
    end 
    return @cell[sheet]["#{row},#{col}"]
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
    if @formula[sheet]["#{row},#{col}"]
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

  # returns all values in this row as an array
  # row numbers are 1,2,3,... like in the spreadsheet
  def row(rownumber,sheet=nil)
    sheet = @default_sheet unless sheet
    read_cells(sheet) unless @cells_read[sheet]
    result = []
    tmp_arr = []
    @cell[sheet].each_pair {|key,value|
      y,x = key.split(',')
      x = x.to_i
      y = y.to_i
      if y == rownumber
        tmp_arr[x] = value
      end
    }
    result = tmp_arr[1..-1]
    while result[-1] == nil
      result = result[0..-2]
    end
    result
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

  # returns all values in this column as an array
  # column numbers are 1,2,3,... like in the spreadsheet
  #--
  #TODO: refactoring nach GenericSpreadsheet?
  def column(columnnumber, sheet=nil)
    if columnnumber.class == String
      columnnumber = GenericSpreadsheet.letter_to_number(columnnumber)
    end
    sheet = @default_sheet unless sheet
    read_cells(sheet) unless @cells_read[sheet]
    result = []
    first_row(sheet).upto(last_row(sheet)) do |row|
      result << cell(row,columnnumber,sheet)
    end
    result
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
    @gs.add_to_cell_roo(row,col,value,sheet_no)
  end

  # returns the first non-empty row in a sheet
  def first_row(sheet=nil)
    sheet = @default_sheet unless sheet
    unless @first_row[sheet]
      sheet_no = sheets.index(sheet) + 1
      @first_row[sheet], @last_row[sheet], @first_column[sheet], @last_column[sheet] = @gs.oben_unten_links_rechts(sheet_no)
    end   
    return @first_row[sheet]
  end

  # returns the last non-empty row in a sheet
  def last_row(sheet=nil)
    sheet = @default_sheet unless sheet
    unless @last_row[sheet]
      sheet_no = sheets.index(sheet) + 1
      @first_row[sheet], @last_row[sheet], @first_column[sheet], @last_column[sheet] = @gs.oben_unten_links_rechts(sheet_no)
    end
    return @last_row[sheet]
  end

  # returns the first non-empty column in a sheet
  def first_column(sheet=nil)
    sheet = @default_sheet unless sheet
    unless @first_column[sheet]
      sheet_no = sheets.index(sheet) + 1
      @first_row[sheet], @last_row[sheet], @first_column[sheet], @last_column[sheet] = @gs.oben_unten_links_rechts(sheet_no)
    end
    return @first_column[sheet]
  end

  # returns the last non-empty column in a sheet
  def last_column(sheet=nil)
    sheet = @default_sheet unless sheet
    unless @last_column[sheet]
      sheet_no = sheets.index(sheet) + 1
      @first_row[sheet], @last_row[sheet], @first_column[sheet], @last_column[sheet] = @gs.oben_unten_links_rechts(sheet_no)
    end
    return @last_column[sheet]
  end

  private

  # read all cells in a sheet
  def read_cells(sheet=nil)
    sheet = @default_sheet unless sheet
    raise RangeError, "illegal sheet <#{sheet}>" unless sheets.index(sheet)
    sheet_no = sheets.index(sheet)+1
    doc = @gs.fulldoc(sheet_no)
    (doc/"gs:cell").each {|item|
      row = item['row']
      col = item['col']
      value = item['inputvalue']
      numericvalue = item['numericvalue']
      if value[0,1] == '='
        formula = value
      else
        formula = nil
      end
      @cell_type[sheet] = {} unless @cell_type[sheet]
      if formula
        ty = :formula
        if numeric?(numericvalue)
          value = numericvalue.to_f
        else
          value = numericvalue
        end
      elsif Google.date?(value)
        ty = :date
      elsif Google.datetime?(value)
        ty = :datetime
      elsif numeric?(value) # or o.class ???
        ty = :float
        value = value.to_f
      elsif Google.time?(value)
        ty = :time
        value = Google.timestring_to_seconds(value)
      else
        ty = :string
      end
      key = "#{row},#{col}"
      @cell[sheet] = {} unless @cell[sheet]
      if ty == :date
        dd,mm,yyyy = value.split('/')
        @cell[sheet][key] = sprintf("%04d-%02d-%02d",yyyy.to_i,mm.to_i,dd.to_i)
      elsif ty == :datetime
        date_part,time_part = value.split(' ')
        dd,mm,yyyy = date_part.split('/')
        hh,mi,ss = time_part.split(':')
        @cell[sheet][key] = sprintf("%04d-%02d-%02d %02d:%02d:%02d",
          yyyy.to_i,mm.to_i,dd.to_i,hh.to_i,mi.to_i,ss.to_i)
      else
        @cell[sheet][key] = value unless value == "" or value == nil
      end
      @cell_type[sheet][key] = ty # Openoffice.oo_type_2_roo_type(vt)
      @formula[sheet] = {} unless @formula[sheet]
      @formula[sheet][key] = formula  if formula
    }
    @cells_read[sheet] = true
  end

  def numeric?(string)
    string =~ /^[0-9]+[\.]*[0-9]*$/
  end

  # convert string DD/MM/YYYY into a Date-object
  #TODO: was ist mit verschiedenen Typen der Datumseingabe bei Google?
  def Google.to_date(string)
    if string.strip =~ /^([0-9]+)\/([0-9]+)\/([0-9]+)$/
      return Date.new($3.to_i,$2.to_i,$1.to_i)
    else
      return nil
    end
  end


end # class
