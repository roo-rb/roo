
require 'rubygems'
require 'rexml/document'
require 'fileutils'
require 'zip/zipfilesystem'
require 'date'
require 'base64'

class Openoffice < GenericSpreadsheet

  @@nr = 0

  # initialization and opening of a spreadsheet file
  # values for packed: :zip
  def initialize(filename, packed=nil, file_warning=:error) #, create = false)
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
      file_type_check(filename,'.ods','an openoffice')
      #if create and ! File.exists?(filename)
      #  self.create_openoffice(filename)
      #end
      @cells_read = Hash.new
      #TODO: @cells_read[:default] = false
      @filename = filename
      unless File.file?(@filename)
        raise IOError, "file #{@filename} does not exist"
      end
      @@nr += 1
      @file_nr = @@nr
      extract_content
      file = File.new(File.join(@tmpdir, @file_nr.to_s+"_roo_content.xml"))
      @doc = REXML::Document.new file
      file.close
    ensure
      #if ENV["roo_local"] != "thomas-p"
      FileUtils::rm_r(@tmpdir)
      #end
    end
    @default_sheet = nil
    # no need to set default_sheet if there is only one sheet in the document
    if self.sheets.size == 1
      @default_sheet = self.sheets.first
    end
    @cell = Hash.new
    @cell_type = Hash.new
    @formula = Hash.new
    @first_row = Hash.new
    @last_row = Hash.new
    @first_column = Hash.new
    @last_column = Hash.new
    @header_line = 1
  end

  # creates a new empty openoffice-spreadsheet file
  def create_openoffice(filename) #:nodoc:
    #TODO: a better way for creating the file contents
    # now you have to call mkbase64...rb to create an include file with all
    # the empty files in an openoffice zip-file
    load 'base64include.rb'
    # puts @@empty_spreadsheet
    f = File.open(filename,'wb')
    f.print(Base64.decode64(@@empty_spreadsheet))
    f.close
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

  # returns an array of sheet names in the spreadsheet
  def sheets
    return_sheets = []
    oo_document_count = 0
    @doc.each_element do |oo_document|
      oo_document_count += 1
      oo_element_count = 0
      oo_document.each_element do |oo_element|
        oo_element_count += 1
        if oo_element.name == "body"
          oo_element.each_element do |be|
            if be.name == "spreadsheet"
              be.each_element do |se|
                if se.name == "table"
                  return_sheets << se.attributes['name']
                end
              end
            end
          end
        end
      end
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
    @doc.each_element do |oo_document|
      @officeversion = oo_document.attributes['version']
    end
  end

  # helper function to set the internal representation of cells
  def set_cell_values(sheet,x,y,i,v,vt,formula,tr,str_v)
    key = [y,x+i]
    @cell_type[sheet] = {} unless @cell_type[sheet]
    @cell_type[sheet][key] = Openoffice.oo_type_2_roo_type(vt)
    @formula[sheet] = {} unless @formula[sheet]
    @formula[sheet][key] = formula  if formula
    @cell[sheet]    = {} unless @cell[sheet]
    case @cell_type[sheet][key]
    when :float
      @cell[sheet][key] = v.to_f
    when :string
      @cell[sheet][key] = str_v
    when :date
      if tr.attributes['date-value'].size != "XXXX-XX-XX".size
        #-- dann ist noch eine Uhrzeit vorhanden
        #-- "1961-11-21T12:17:18"
        @cell[sheet][key] = DateTime.parse(tr.attributes['date-value'])
        @cell_type[sheet][key] = :datetime
      else
        @cell[sheet][key] = tr.attributes['date-value']
      end
    when :percentage
      @cell[sheet][key] = v.to_f
    when :time
      hms = v.split(':')
      @cell[sheet][key] = hms[0].to_i*3600 + hms[1].to_i*60 + hms[2].to_i
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
    oo_document_count = 0
    @doc.each_element do |oo_document|
      # @officeversion = oo_document.attributes['version']
      oo_document_count += 1
      oo_element_count = 0
      oo_document.each_element do |oo_element|
        oo_element_count += 1
        if oo_element.name == "body"
          oo_element.each_element do |be|
            if be.name == "spreadsheet"
              be.each_element do |se|
                if se.name == "table"
                  if se.attributes['name']==sheet
                    sheet_found = true
                    x=1
                    y=1
                    se.each_element do |te|
                      if te.name == "table-column"
                        rep = te.attributes["number-columns-repeated"]
                      elsif te.name == "table-row"
                        if te.attributes['number-rows-repeated']
                          skip_y = te.attributes['number-rows-repeated'].to_i
                          y = y + skip_y - 1 # minus 1 because this line will be counted as a line element
                        end
                        te.each_element do |tr|
                          if tr.name == 'table-cell'
                            skip = tr.attributes['number-columns-repeated']
                            formula = tr.attributes['formula']
                            vt = tr.attributes['value-type']
                            v  = tr.attributes['value']
                            if vt == 'string'
                              str_v  = ''
                              # insert \n if there is more than one paragraph
                              para_count = 0
                              tr.each_element do |str|
                                if str.name == 'p'
                                  v = str.text
                                  str_v += "\n" if para_count > 0
                                  para_count += 1
                                  if str.children.size > 1
                                    str_v = children_to_string(str.children)
                                  else
                                    str.children.each {|child|
                                      str_v = str_v + child.to_s #.text
                                    }
                                  end
                                  str_v.gsub!(/&quot;/,'"')
                                  str_v.gsub!(/&amp;/,'&')
                                  str_v.gsub!(/&apos;/,"'")
                                end # == 'p'
                              end
                            elsif vt == 'time'
                              tr.each_element do |str|
                                if str.name == 'p'
                                  v = str.text
                                end
                              end
                            elsif vt == '' or vt == nil
                              #
                            elsif vt == 'date'
                              #
                            elsif vt == 'percentage'
                              #
                            elsif vt == 'float'
                              #
                            elsif vt == 'boolean'
                              #
                            else
                              # raise "unknown type #{vt}"
                            end
                            if skip
                              if v != nil or tr.attributes['date-value']
                                0.upto(skip.to_i-1) do |i|
                                  set_cell_values(sheet,x,y,i,v,vt,formula,tr,str_v)
                                end
                              end
                              x += (skip.to_i - 1)
                            end # if skip
                            set_cell_values(sheet,x,y,0,v,vt,formula,tr,str_v)
                            x += 1
                          end
                        end
                        y += 1
                        x = 1
                      end
                    end
                  end # sheet
                end
              end
            end
          end
        end
      end
    end
    if !sheet_found
      raise RangeError
    end
    @cells_read[sheet] = true
  end

  # Checks if the default_sheet exists. If not an RangeError exception is
  # raised
  def check_default_sheet
    sheet_found = false
    raise ArgumentError, "Error: default_sheet not set" if @default_sheet == nil
    @doc.each_element do |oo_document|
      oo_element_count = 0
      oo_document.each_element do |oo_element|
        oo_element_count += 1
        if oo_element.name == "body"
          oo_element.each_element do |be|
            if be.name == "spreadsheet"
              be.each_element do |se|
                if se.name == "table"
                  if se.attributes['name'] == @default_sheet
                    sheet_found = true
                  end # sheet
                end
              end
            end
          end
        end
      end
    end
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

  def Openoffice.oo_type_2_roo_type(ootype)
    return A_ROO_TYPE[ootype]
  end
 
  # helper method to convert compressed spaces and other elements within
  # an text into a string
  def children_to_string(children)
    result = ''
    children.each {|child|
      if child.class == REXML::Text
        result = result + child.to_s
      else
        if child.name == 's'
          compressed_spaces = child.attributes['c'].to_i
          # no explicit number means a count of 1:
          if compressed_spaces == 0
            compressed_spaces = 1
          end
          result = result + " "*compressed_spaces
        else
          result = result + child.to_s
        end
      end
    }
    result
  end

end # class
