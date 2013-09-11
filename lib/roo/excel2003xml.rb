require 'date'
require 'base64'
require 'nokogiri'

class Roo::Excel2003XML < Roo::Base

  # initialization and opening of a spreadsheet file
  # values for packed: :zip
  def initialize(filename, options={}, deprecated_file_warning=:error)
    if Hash === options
      packed = options[:packed]
      file_warning = options[:file_warning] || :error
    else
      warn 'Supplying `packed` or `file_warning` as separate arguments to `Roo::Excel2003XML.new` is deprecated. Use an options hash instead.'
      packed = options
      file_warning = deprecated_file_warning
    end

    make_tmpdir do |tmpdir|
      filename = download_uri(filename, tmpdir) if uri?(filename)
      filename = unzip(filename, tmpdir) if packed == :zip

      file_type_check(filename,'.xml','an Excel 2003 XML', file_warning)
      @filename = filename
      unless File.file?(@filename)
        raise IOError, "file #{@filename} does not exist"
      end
      @doc = load_xml(@filename)
    end
    super(filename, options)
    @formula = Hash.new
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
    sheet ||= @default_sheet
    read_cells(sheet)
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
    sheet ||= @default_sheet
    read_cells(sheet)
    row,col = normalize(row,col)
    @formula[sheet][[row,col]] && @formula[sheet][[row,col]]["oooc:".length..-1]
  end
  alias_method :formula?, :formula

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
    sheet ||= @default_sheet
    read_cells(sheet)
    row,col = normalize(row,col)
    style_name = @style[sheet][[row,col]] || @style_defaults[sheet][col - 1] || 'Default'
    @style_definitions[style_name]
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
    sheet ||= @default_sheet
    read_cells(sheet)
    row,col = normalize(row,col)
    if @formula[sheet][[row,col]]
      return :formula
    else
      @cell_type[sheet][[row,col]]
    end
  end

  def sheets
    @doc.xpath("/ss:Workbook/ss:Worksheet").map do |sheet|
      sheet['ss:Name']
    end
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
    sheet ||= @default_sheet
    read_cells(sheet)
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
    sheet ||= @default_sheet
    read_cells(sheet)
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
      @officeversion = office['version']
    end
  end

  # helper function to set the internal representation of cells
  def set_cell_values(sheet,x,y,i,v,value_type,formula,table_cell,str_v,style_name)
    key = [y,x+i]
    @cell_type[sheet] = {} unless @cell_type[sheet]
    @cell_type[sheet][key] = value_type
    @formula[sheet] = {} unless @formula[sheet]
    @formula[sheet][key] = formula  if formula
    @cell[sheet]    = {} unless @cell[sheet]
    @style[sheet] = {} unless @style[sheet]
    @style[sheet][key] = style_name
    @cell[sheet][key] =
      case @cell_type[sheet][key]
      when :float
        v.to_f
      when :string
        str_v
      when :datetime
        DateTime.parse(v)
      when :percentage
        v.to_f
      # when :time
      #   hms = v.split(':')
      #   hms[0].to_i*3600 + hms[1].to_i*60 + hms[2].to_i
      else
        v
      end
  end

  # read all cells in the selected sheet
  #--
  # the following construct means '4 blanks'
  # some content <text:s text:c="3"/>
  #++
  def read_cells(sheet=nil)
    sheet ||= @default_sheet
    validate_sheet!(sheet)
    return if @cells_read[sheet]
    sheet_found = false
    @doc.xpath("/ss:Workbook/ss:Worksheet[@ss:Name='#{sheet}']").each do |ws|
      sheet_found = true
      row = 1
      col = 1
      column_attributes = {}
      idx = 0
      ws.xpath('./ss:Table/ss:Column').each do |c|
        column_attributes[(idx += 1).to_s] = c['StyleID']
      end
      ws.xpath('./ss:Table/ss:Row').each do |r|
        skip_to_row = r['Index'].to_i
        row = skip_to_row if skip_to_row > 0
        style_name = r['StyleID'] if r['StyleID']
        r.xpath('./ss:Cell').each do |c|
          skip_to_col = c['Index'].to_i
          col = skip_to_col if skip_to_col > 0
          if c['StyleID']
            style_name = c['StyleID']
          elsif
            style_name ||= column_attributes[c['Index']]
          end
          c.xpath('./ss:Data').each do |cell|
            formula = cell['Formula']
            value_type = cell['ss:Type'].downcase.to_sym
            v =  cell.content
            str_v = v
            case value_type
            when :number
              v = v.to_f
              value_type = :float
            when :datetime
              if v =~ /^1899-12-31T(\d{2}:\d{2}:\d{2})/
                v = $1
                value_type = :time
              elsif v =~ /([^T]+)T00:00:00.000/
                v = $1
                value_type = :date
              end
            when :boolean
              v = cell['boolean-value']
            end
            set_cell_values(sheet,col,row,0,v,value_type,formula,cell,str_v,style_name)
          end
          col += 1
        end
        row += 1
        col = 1
      end
    end
    if !sheet_found
      raise RangeError, "Unable to find sheet #{sheet} for reading"
    end
    @cells_read[sheet] = true
  end

  def read_styles
    @doc.xpath("/ss:Workbook/ss:Styles/ss:Style").each do |style|
      style_id = style['ID']
      @style_definitions[style_id] = Roo::Excel2003XML::Font.new
      if font = style.at_xpath('./ss:Font')
        @style_definitions[style_id].bold = font['Bold']
        @style_definitions[style_id].italic = font['Italic']
        @style_definitions[style_id].underline = font['Underline']
      end
    end
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
          compressed_spaces = child['c'].to_i
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
