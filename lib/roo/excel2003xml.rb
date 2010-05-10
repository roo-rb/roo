require 'xml'
require 'fileutils'
require 'zip/zipfilesystem'
require 'date'
require 'base64'
require 'cgi'

class Roo::Excel2003XML < Roo::GenericSpreadsheet

  @@nr = 0

  # initialization and opening of a spreadsheet file
  # values for packed: :zip
  def initialize(filename, packed=nil, file_warning=:error) 
    @file_warning = file_warning
    super()
    @tmpdir = "oo_"+$$.to_s
    @tmpdir = File.join(ENV['ROO_TMP'], @tmpdir) if ENV['ROO_TMP'] 
    unless File.exists?(@tmpdir)
      FileUtils::mkdir(@tmpdir)
    end
    filename = open_from_uri(filename) if filename[0,7] == "http://"
    filename = unzip(filename) if packed and packed == :zip
    begin
      file_type_check(filename,'.xml','an Excel 2003 XML')
      @cells_read = Hash.new
      @filename = filename
      unless File.file?(@filename)
        raise IOError, "file #{@filename} does not exist"
      end
      @doc = XML::Parser.file(@filename).parse
    ensure
      FileUtils::rm_r(@tmpdir)
    end
    @default_sheet = self.sheets.first
    @cell = Hash.new
    @cell_type = Hash.new
    @formula = Hash.new
    @first_row = Hash.new
    @last_row = Hash.new
    @first_column = Hash.new
    @last_column = Hash.new
    @header_line = 1
    @style = Hash.new
    @style_defaults = Hash.new { |h,k| h[k] = [] }
    @style_definitions = Hash.new 
    read_styles
  end

  # Returns the content of a spreadsheet-cell.
  # (1,1) is the upper left corner.
  # (1,1), (1,'A'), ('A',1), ('a',1) all refers to the
  # cell at the first line and first row.
  def cell(row, col, sheet=nil)
    sheet = @default_sheet unless sheet
    read_cells(sheet) unless @cells_read[sheet]
    row,col = normalize(row,col)
    if celltype(row,col,sheet) == :date
      yyyy,mm,dd = @cell[sheet][[row,col]].split('-')
      return Date.new(yyyy.to_i,mm.to_i,dd.to_i)
    end
    @cell[sheet][[row,col]]
  end

  # Returns the formula at (row,col).
  # Returns nil if there is no formula.
  # The method #formula? checks if there is a formula.
  def formula(row,col,sheet=nil)
    sheet = @default_sheet unless sheet
    read_cells(sheet) unless @cells_read[sheet]
    row,col = normalize(row,col)
    if @formula[sheet][[row,col]] == nil
      return nil
    else
      return @formula[sheet][[row,col]]["oooc:".length..-1]
    end
  end

  # true, if there is a formula
  def formula?(row,col,sheet=nil)
    sheet = @default_sheet unless sheet
    read_cells(sheet) unless @cells_read[sheet]
    row,col = normalize(row,col)
    formula(row,col) != nil
  end
  
  class Font
    attr_accessor :bold, :italic, :underline
    
    def bold? 
      @bold == '1'
    end

    def italic? 
      @italic == '1'
    end
    
    def underline? 
      @underline != nil
    end
  end

  # Given a cell, return the cell's style 
  def font(row, col, sheet=nil)
    sheet = @default_sheet unless sheet
    read_cells(sheet) unless @cells_read[sheet]
    row,col = normalize(row,col)
    style_name = @style[sheet][[row,col]] || @style_defaults[sheet][col - 1] || 'Default'
    @style_definitions[style_name]
  end 
  
  # set a cell to a certain value
  # (this will not be saved back to the spreadsheet file!)
  def set(row,col,value,sheet=nil) #:nodoc:
    sheet = @default_sheet unless sheet
    read_cells(sheet) unless @cells_read[sheet]
    row,col = normalize(row,col)
    set_value(row,col,value,sheet)
    if value.class == Fixnum
      set_type(row,col,:float,sheet)
    elsif value.class == String
      set_type(row,col,:string,sheet)
    elsif value.class == Float
      set_type(row,col,:string,sheet)
    else
      raise ArgumentError, "Type for "+value.to_s+" not set"
    end
  end

  # returns the type of a cell:
  # * :float
  # * :string
  # * :date
  # * :percentage
  # * :formula
  # * :time
  # * :datetime
  def celltype(row,col,sheet=nil)
    sheet = @default_sheet unless sheet
    read_cells(sheet) unless @cells_read[sheet]
    row,col = normalize(row,col)
    if @formula[sheet][[row,col]]
      return :formula
    else
      @cell_type[sheet][[row,col]]
    end
  end

  def sheets
    return_sheets = []
    @doc.find("//ss:Worksheet").each do |sheet|
      return_sheets << sheet.attributes['Name']
    end
    return_sheets
  end
    
  # version of the openoffice document
  # at 2007 this is always "1.0"
  def officeversion
    oo_version
    @officeversion
  end

  # shows the internal representation of all cells
  # mainly for debugging purposes
  def to_s(sheet=nil)
    sheet = @default_sheet unless sheet
    read_cells(sheet) unless @cells_read[sheet]
    @cell[sheet].inspect
  end

  # save spreadsheet
  def save #:nodoc:
    42
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

  private

  # read the version of the OO-Version
  def oo_version
    @doc.find("//*[local-name()='document-content']").each do |office|
      @officeversion = office.attributes['version']
    end
  end

  # helper function to set the internal representation of cells
  def set_cell_values(sheet,x,y,i,v,vt,formula,table_cell,str_v,style_name)
    key = [y,x+i]    
    @cell_type[sheet] = {} unless @cell_type[sheet]
    @cell_type[sheet][key] = vt
    @formula[sheet] = {} unless @formula[sheet]
    @formula[sheet][key] = formula  if formula
    @cell[sheet]    = {} unless @cell[sheet]
    @style[sheet] = {} unless @style[sheet]
    @style[sheet][key] = style_name
    case @cell_type[sheet][key]
    when :float
      @cell[sheet][key] = v.to_f
    when :string
      @cell[sheet][key] = str_v
    when :datetime
      @cell[sheet][key] = DateTime.parse(v)
      @cell_type[sheet][key] = :datetime
    when :percentage
      @cell[sheet][key] = v.to_f
    # when :time
    #   hms = v.split(':')
    #   @cell[sheet][key] = hms[0].to_i*3600 + hms[1].to_i*60 + hms[2].to_i
    else
      @cell[sheet][key] = v
    end
  end

  # read all cells in the selected sheet
  #--
  # the following construct means '4 blanks'
  # some content <text:s text:c="3"/>
  #++
  def read_cells(sheet=nil)
    sheet = @default_sheet unless sheet
    sheet_found = false
    raise ArgumentError, "Error: sheet '#{sheet||'nil'}' not valid" if @default_sheet == nil and sheet==nil
    raise RangeError unless self.sheets.include? sheet
    @doc.find("ss:Worksheet[@ss:Name='#{sheet}']").each do |ws|
        sheet_found = true
        row = 1
        col = 1
        column_attributes = {}   
        idx = 0
        ws.find('.//ss:Column').each do |c|
          column_attributes[(idx += 1).to_s] = c.attributes['StyleID'] 
        end  
        ws.find('.//ss:Row').each do |r|
          skip_to_row = r.attributes['Index'].to_i
          row = skip_to_row if skip_to_row > 0
          style_name = r.attributes['StyleID'] if r.attributes['StyleID'] 
          r.each do |c|    
            next unless c.name == 'Cell'        
            skip_to_col = c.attributes['Index'].to_i
            col = skip_to_col if skip_to_col > 0
            if c.attributes['StyleID'] 
              style_name = c.attributes['StyleID'] 
            elsif 
              style_name ||= column_attributes[c.attributes['Index']]
            end  
            c.each_element do |cell|
              formula = nil
              if cell.name == 'Data'
                formula = cell.attributes['Formula']
                vt = cell.attributes['Type'].downcase.to_sym
                v =  cell.content
                str_v = v
                case vt 
                when :number
                  v = v.to_f   
                  vt = :float
                when :datetime
                  if v =~ /^1899-12-31T(\d{2}:\d{2}:\d{2})/
                    v = $1
                    vt = :time
                  elsif v =~ /([^T]+)T00:00:00.000/
                    v = $1
                    vt = :date
                  end
                when :boolean
                  v = cell.attributes['boolean-value']
                end
              end
              set_cell_values(sheet,col,row,0,v,vt,formula,cell,str_v,style_name)
            end    
            col += 1 
          end
          row += 1
          col = 1
        end    
    end
    if !sheet_found
      raise RangeError
    end
    @cells_read[sheet] = true
  end

  def read_styles
    @doc.find("ss:Styles").each do |styles|
       styles.find('.//ss:Style').each do |style|
         style_id = style.attributes['ID']
         @style_definitions[style_id] = Roo::Excel2003XML::Font.new
         font = style.find_first('.//ss:Font')
         if font
           @style_definitions[style_id].bold = font.attributes['Bold']
           @style_definitions[style_id].italic = font.attributes['Italic']
           @style_definitions[style_id].underline = font.attributes['Underline']
        end
      end
    end  
  end
  
  # Checks if the default_sheet exists. If not an RangeError exception is
  # raised
  def check_default_sheet
    sheet_found = false
    raise ArgumentError, "Error: default_sheet not set" if @default_sheet == nil
    sheet_found = true if sheets.include?(@default_sheet)
    if ! sheet_found
      raise RangeError, "sheet '#{@default_sheet}' not found"
    end
  end

  def process_zipfile(zip, path='')
    if zip.file.file? path
      if path == "content.xml"
        open(File.join(@tmpdir, @file_nr.to_s+'_roo_content.xml'),'wb') {|f|
          f << zip.read(path)
        }
      end
    else
      unless path.empty?
        path += '/'
      end
      zip.dir.foreach(path) do |filename|
        process_zipfile(zip, path+filename)
      end
    end
  end

  def extract_content
    Zip::ZipFile.open(@filename) do |zip|
      process_zipfile(zip)
    end
  end

  def set_value(row,col,value,sheet=nil)
    sheet = @default_value unless sheet
    @cell[sheet][[row,col]] = value
  end

  def set_type(row,col,type,sheet=nil)
    sheet = @default_value unless sheet
    @cell_type[sheet][[row,col]] = type
  end

  A_ROO_TYPE = {
    "float"      => :float,
    "string"     => :string,
    "date"       => :date,
    "percentage" => :percentage,
    "time"       => :time,
  }

  def self.oo_type_2_roo_type(ootype)
    return A_ROO_TYPE[ootype]
  end
 
  # helper method to convert compressed spaces and other elements within
  # an text into a string
  def children_to_string(children)
    result = ''
    children.each {|child|
      if child.text?
        result = result + child.content
      else
        if child.name == 's'
          compressed_spaces = child.attributes['c'].to_i
          # no explicit number means a count of 1:
          if compressed_spaces == 0
            compressed_spaces = 1
          end
          result = result + " "*compressed_spaces
        else
          result = result + child.content
        end
      end
    }
    result
  end

end # class
