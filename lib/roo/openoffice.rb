require 'date'
require 'nokogiri'
require 'cgi'

class Roo::OpenOffice < Roo::Base
  class << self
    def extract_content(tmpdir, filename)
      Roo::ZipFile.open(filename) do |zip|
        process_zipfile(tmpdir, zip)
      end
    end

    def process_zipfile(tmpdir, zip, path='')
      if zip.file.file? path
        if path == "content.xml"
          open(File.join(tmpdir, 'roo_content.xml'),'wb') {|f|
            f << zip.read(path)
          }
        end
      else
        unless path.empty?
          path += '/'
        end
        zip.dir.foreach(path) do |filename|
          process_zipfile(tmpdir, zip, path+filename)
        end
      end
    end
  end

  # initialization and opening of a spreadsheet file
  # values for packed: :zip
  def initialize(filename, options={}, deprecated_file_warning=:error, deprecated_tmpdir_root=nil)
    if Hash === options
      packed = options[:packed]
      file_warning = options[:file_warning] || :error
      tmpdir_root = options[:tmpdir_root]
    else
      warn 'Supplying `packed`, `file_warning`, or `tmpdir_root` as separate arguments to `Roo::OpenOffice.new` is deprecated. Use an options hash instead.'
      packed = options
      file_warning = deprecated_file_warning
      tmpdir_root = deprecated_tmpdir_root
    end

    file_type_check(filename,'.ods','an Roo::OpenOffice', file_warning, packed)
    make_tmpdir(tmpdir_root) do |tmpdir|
      filename = download_uri(filename, tmpdir) if uri?(filename)
      filename = unzip(filename, tmpdir) if packed == :zip
      #TODO: @cells_read[:default] = false
      @filename = filename
      unless File.file?(@filename)
        raise IOError, "file #{@filename} does not exist"
      end
      self.class.extract_content(tmpdir, @filename)
      @doc = load_xml(File.join(tmpdir, "roo_content.xml"))
    end
    super(filename, options)
    @formula = Hash.new
    @style = Hash.new
    @style_defaults = Hash.new { |h,k| h[k] = [] }
    @style_definitions = Hash.new
    @comment = Hash.new
    @comments_read = Hash.new
  end

  def method_missing(m,*args)
    read_labels
    # is method name a label name
	  if @label.has_key?(m.to_s)
      row,col = label(m.to_s)
		  cell(row,col)
	  else
		  # call super for methods like #a1
		  super
	  end
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
      yyyy,mm,dd = @cell[sheet][[row,col]].to_s.split('-')
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
    @formula[sheet][[row,col]]
  end
  alias_method :formula?, :formula

  # returns each formula in the selected sheet as an array of elements
  # [row, col, formula]
  def formulas(sheet=nil)
    sheet ||= @default_sheet
    read_cells(sheet)
    if @formula[sheet]
      @formula[sheet].each.collect do |elem|
        [elem[0][0], elem[0][1], elem[1]]
      end
    else
      []
    end
  end

  class Font
    attr_accessor :bold, :italic, :underline

    def bold?
      @bold == 'bold'
    end

    def italic?
      @italic == 'italic'
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
    @doc.xpath("//*[local-name()='table']").map do |sheet|
      sheet.attributes["name"].value
    end
  end

  # version of the Roo::OpenOffice document
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

  # returns the row,col values of the labelled cell
  # (nil,nil) if label is not defined
  def label(labelname)
    read_labels
    unless @label.size > 0
      return nil,nil,nil
    end
    if @label.has_key? labelname
      return @label[labelname][1].to_i,
        Roo::Base.letter_to_number(@label[labelname][2]),
        @label[labelname][0]
    else
      return nil,nil,nil
    end
  end

  # Returns an array which all labels. Each element is an array with
  # [labelname, [row,col,sheetname]]
  def labels(sheet=nil)
    read_labels
    @label.map do |label|
      [ label[0], # name
        [ label[1][1].to_i, # row
          Roo::Base.letter_to_number(label[1][2]), # column
          label[1][0], # sheet
        ] ]
    end
  end

  # returns the comment at (row/col)
  # nil if there is no comment
  def comment(row,col,sheet=nil)
    sheet ||= @default_sheet
    read_cells(sheet)
    row,col = normalize(row,col)
    return nil unless @comment[sheet]
    @comment[sheet][[row,col]]
  end

  # true, if there is a comment
  def comment?(row,col,sheet=nil)
    sheet ||= @default_sheet
    read_cells(sheet)
    row,col = normalize(row,col)
    comment(row,col) != nil
  end


  # returns each comment in the selected sheet as an array of elements
  # [row, col, comment]
  def comments(sheet=nil)
    sheet ||= @default_sheet
    read_comments(sheet) unless @comments_read[sheet]
    if @comment[sheet]
      @comment[sheet].each.collect do |elem|
        [elem[0][0],elem[0][1],elem[1]]
      end
    else
      []
    end
  end

  private

  # read the version of the OO-Version
  def oo_version
    @doc.xpath("//*[local-name()='document-content']").each do |office|
      @officeversion = attr(office,'version')
    end
  end

  # helper function to set the internal representation of cells
  def set_cell_values(sheet,x,y,i,v,value_type,formula,table_cell,str_v,style_name)
    key = [y,x+i]
    @cell_type[sheet] = {} unless @cell_type[sheet]
    @cell_type[sheet][key] = Roo::OpenOffice.oo_type_2_roo_type(value_type)
    @formula[sheet] = {} unless @formula[sheet]
    if formula
      ['of:', 'oooc:'].each do |prefix|
        if formula[0,prefix.length] == prefix
          formula = formula[prefix.length..-1]
        end
      end
      @formula[sheet][key] = formula
    end
    @cell[sheet] = {} unless @cell[sheet]
    @style[sheet] = {} unless @style[sheet]
    @style[sheet][key] = style_name
    case @cell_type[sheet][key]
    when :float
      @cell[sheet][key] = v.to_f
    when :string
      @cell[sheet][key] = str_v
    when :date
      #TODO: if table_cell.attributes['date-value'].size != "XXXX-XX-XX".size
      if attr(table_cell,'date-value').size != "XXXX-XX-XX".size
        #-- dann ist noch eine Uhrzeit vorhanden
        #-- "1961-11-21T12:17:18"
        @cell[sheet][key] = DateTime.parse(attr(table_cell,'date-value').to_s)
        @cell_type[sheet][key] = :datetime
      else
        @cell[sheet][key] = table_cell.attributes['date-value']
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
    sheet ||= @default_sheet
    validate_sheet!(sheet)
    return if @cells_read[sheet]

    sheet_found = false
    @doc.xpath("//*[local-name()='table']").each do |ws|
      if sheet == attr(ws,'name')
        sheet_found = true
        col = 1
        row = 1
        ws.children.each do |table_element|
          case table_element.name
          when 'table-column'
            @style_defaults[sheet] << table_element.attributes['default-cell-style-name']
          when 'table-row'
            if table_element.attributes['number-rows-repeated']
              skip_row = attr(table_element,'number-rows-repeated').to_s.to_i
              row = row + skip_row - 1
            end
            table_element.children.each do |cell|
              skip_col = attr(cell, 'number-columns-repeated')
              formula = attr(cell,'formula')
              value_type = attr(cell,'value-type')
              v =  attr(cell,'value')
              style_name = attr(cell,'style-name')
              case value_type
              when 'string'
                str_v  = ''
                # insert \n if there is more than one paragraph
                para_count = 0
                cell.children.each do |str|
                  # begin comments
=begin
- <table:table-cell office:value-type="string">
  - <office:annotation office:display="true" draw:style-name="gr1" draw:text-style-name="P1" svg:width="1.1413in" svg:height="0.3902in" svg:x="2.0142in" svg:y="0in" draw:caption-point-x="-0.2402in" draw:caption-point-y="0.5661in">
      <dc:date>2011-09-20T00:00:00</dc:date>
      <text:p text:style-name="P1">Kommentar fuer B4</text:p>
    </office:annotation>
    <text:p>B4 (mit Kommentar)</text:p>
  </table:table-cell>
=end
                  if str.name == 'annotation'
                    str.children.each do |annotation|
                      if annotation.name == 'p'
                        # @comment ist ein Hash mit Sheet als Key (wie bei @cell)
                        # innerhalb eines Elements besteht ein Eintrag aus einem
                        # weiteren Hash mit Key [row,col] und dem eigentlichen
                        # Kommentartext als Inhalt
                        @comment[sheet] = Hash.new unless @comment[sheet]
                        key = [row,col]
                        @comment[sheet][key] = annotation.text
                      end
                    end
                  end
                  # end comments
                  if str.name == 'p'
                    v = str.content
                    str_v += "\n" if para_count > 0
                    para_count += 1
                    if str.children.size > 1
                      str_v += children_to_string(str.children)
                    else
                      str.children.each do |child|
                        str_v += child.content #.text
                      end
                    end
                    str_v.gsub!(/&apos;/,"'")  # special case not supported by unescapeHTML
                    str_v = CGI.unescapeHTML(str_v)
                  end # == 'p'
                end
              when 'time'
                cell.children.each do |str|
                  if str.name == 'p'
                    v = str.content
                  end
                end
              when '', nil
                #
              when 'date'
                #
              when 'percentage'
                #
              when 'float'
                #
              when 'boolean'
                v = attr(cell,'boolean-value').to_s
              else
                # raise "unknown type #{value_type}"
              end
              if skip_col
                if v != nil or cell.attributes['date-value']
                  0.upto(skip_col.to_i-1) do |i|
                    set_cell_values(sheet,col,row,i,v,value_type,formula,cell,str_v,style_name)
                  end
                end
                col += (skip_col.to_i - 1)
              end # if skip
              set_cell_values(sheet,col,row,0,v,value_type,formula,cell,str_v,style_name)
              col += 1
            end
            row += 1
            col = 1
          end
        end
      end
    end
    @doc.xpath("//*[local-name()='automatic-styles']").each do |style|
      read_styles(style)
    end
    if !sheet_found
      raise RangeError
    end
    @cells_read[sheet] = true
    @comments_read[sheet] = true
  end

  # Only calls read_cells because Roo::Base calls read_comments
  # whereas the reading of comments is done in read_cells for Roo::OpenOffice-objects
  def read_comments(sheet=nil)
    read_cells(sheet)
  end

  def read_labels
    @label ||= Hash[@doc.xpath("//table:named-range").map do |ne|
      #-
      # $Sheet1.$C$5
      #+
      name = attr(ne,'name').to_s
      sheetname,coords = attr(ne,'cell-range-address').to_s.split('.$')
      col, row = coords.split('$')
      sheetname = sheetname[1..-1] if sheetname[0,1] == '$'
      [name, [sheetname,row,col]]
    end]
  end

  def read_styles(style_elements)
    @style_definitions['Default'] = Roo::OpenOffice::Font.new
    style_elements.each do |style|
      next unless style.name == 'style'
      style_name = attr(style,'name')
      style.each do |properties|
        font = Roo::OpenOffice::Font.new
        font.bold = attr(properties,'font-weight')
        font.italic = attr(properties,'font-style')
        font.underline = attr(properties,'text-underline-style')
        @style_definitions[style_name] = font
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
          compressed_spaces = child.attributes['c'].to_s.to_i
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

  def attr(node, attr_name)
    if node.attributes[attr_name]
      node.attributes[attr_name].value
    end
  end
end # class

# LibreOffice is just an alias for Roo::OpenOffice class
class Roo::LibreOffice < Roo::OpenOffice
end
