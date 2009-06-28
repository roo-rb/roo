require 'gdata/spreadsheet'
require 'xml'

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
    @default_sheet = nil
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
    @gs = GData::Spreadsheet.new(spreadsheetkey)
    @gs.authenticate(user, password)
    @sheetlist = @gs.sheetlist
    #-- ----------------------------------------------------------------------
    #-- TODO: Behandlung von Berechtigungen hier noch einbauen ???
    #-- ----------------------------------------------------------------------
    @default_sheet = self.sheets.first
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
    @gs.add_to_cell_roo(row,col,value,sheet_no)
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

  # read all cells in a sheet. 
  def read_cells(sheet=nil)
    sheet = @default_sheet unless sheet
    raise RangeError, "illegal sheet <#{sheet}>" unless sheets.index(sheet)
    sheet_no = sheets.index(sheet)+1
    xml = @gs.fulldoc(sheet_no).to_s
    doc = XML::Parser.string(xml).parse
    doc.find("//*[local-name()='cell']").each do |item|
      row = item['row']
      col = item['col']
      key = "#{row},#{col}"
      string_value =  item['inputvalue'] ||  item['inputValue'] 
      numeric_value = item['numericvalue']  ||  item['numericValue'] 
      (value, value_type) = determine_datatype(string_value, numeric_value)
      @cell[sheet][key] = value unless value == "" or value == nil
      @cell_type[sheet][key] = value_type 
      @formula[sheet] = {} unless @formula[sheet]
      @formula[sheet][key] = string_value if value_type == :formula
    end
    @cells_read[sheet] = true
  end
  
  def determine_datatype(val, numval=nil)
    if val[0,1] == '='
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
    return val, ty 
  end
  
end # class
