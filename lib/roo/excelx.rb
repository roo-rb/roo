require 'fileutils'
require 'zip/zipfilesystem'
require 'date'
require 'rubygems'
require 'nokogiri'

if RUBY_VERSION < '1.9.0'
  class  String
    def end_with?(str)
      self[-str.length,str.length] == str
    end
  end
end

class Roo::Excelx < Roo::GenericSpreadsheet
  FORMATS = {
    'General' => :float,
    '0' => :float,
    '0.00' => :float,
    '#,##0' => :float,
    '#,##0.00' => :float,
    '0%' => :percentage,
    '0.00%' => :percentage,
    '0.00E+00' => :float,
    '# ?/?' => :float, #??? TODO:
    '# ??/??' => :float, #??? TODO:
    'mm-dd-yy' => :date,
    'd-mmm-yy' => :date,
    'd-mmm' => :date,
    'mmm-yy' => :date,
    'h:mm AM/PM' => :date,
    'h:mm:ss AM/PM' => :date,
    'h:mm' => :time,
    'h:mm:ss' => :time,
    'm/d/yy h:mm' => :date,
    '#,##0 ;(#,##0)' => :float,
    '#,##0 ;[Red](#,##0)' => :float,
    '#,##0.00;(#,##0.00)' => :float,
    '#,##0.00;[Red](#,##0.00)' => :float,
    'mm:ss' => :time,
    '[h]:mm:ss' => :time,
    'mmss.0' => :time,
    '##0.0E+0' => :float,
    '@' => :float,
    #-- zusaetzliche Formate, die nicht standardmaessig definiert sind:
    "yyyy\\-mm\\-dd" => :date,
    'dd/mm/yy' => :date,
    'hh:mm:ss' => :time,
    "dd/mm/yy\\ hh:mm" => :datetime,
    'dd/mmm/yy' => :date, # 2011-05-21
    'yyyy-mm-dd' => :date, # 2011-09-16
    # was used in a spreadsheet file from a windows phone
  }
  STANDARD_FORMATS = {
    0 => 'General',
    1 => '0',
    2 => '0.00',
    3 => '#,##0',
    4 => '#,##0.00',
    9 => '0%',
    10 => '0.00%',
    11 => '0.00E+00',
    12 => '# ?/?',
    13 => '# ??/??',
    14 => 'mm-dd-yy',
    15 => 'd-mmm-yy',
    16 => 'd-mmm',
    17 => 'mmm-yy',
    18 => 'h:mm AM/PM',
    19 => 'h:mm:ss AM/PM',
    20 => 'h:mm',
    21 => 'h:mm:ss',
    22 => 'm/d/yy h:mm',
    37 => '#,##0 ;(#,##0)',
    38 => '#,##0 ;[Red](#,##0)',
    39 => '#,##0.00;(#,##0.00)',
    40 => '#,##0.00;[Red](#,##0.00)',
    45 => 'mm:ss',
    46 => '[h]:mm:ss',
    47 => 'mmss.0',
    48 => '##0.0E+0',
    49 => '@',
  }

  # initialization and opening of a spreadsheet file
  # values for packed: :zip
  def initialize(filename, packed=nil, file_warning = :error) #, create = false)
    file_type_check(filename,'.xlsx','an Excel-xlsx', file_warning, packed)
    make_tmpdir do |tmpdir|
      filename = open_from_uri(filename, tmpdir) if uri?(filename)
      filename = unzip(filename, tmpdir) if packed == :zip
      @cells_read = Hash.new
      @filename = filename
      unless File.file?(@filename)
        raise IOError, "file #{@filename} does not exist"
      end
      @comments_files = Array.new
      extract_content(tmpdir, @filename)
      @workbook_doc = File.open(File.join(tmpdir, "roo_workbook.xml")) do |file|
        Nokogiri::XML(file)
      end
      @shared_table = []
      if File.exist?(File.join(tmpdir, 'roo_sharedStrings.xml'))
        @sharedstring_doc = File.open(File.join(tmpdir, 'roo_sharedStrings.xml')) do |file|
          Nokogiri::XML(file)
        end
        read_shared_strings(@sharedstring_doc)
      end
      @styles_table = []
      @style_definitions = Array.new # TODO: ??? { |h,k| h[k] = {} }
      if File.exist?(File.join(tmpdir, 'roo_styles.xml'))
        @styles_doc = File.open(File.join(tmpdir, 'roo_styles.xml')) do |file|
          Nokogiri::XML(file)
        end
        read_styles(@styles_doc)
      end
      @sheet_doc = @sheet_files.map do |item|
        File.open(item) do |file|
          Nokogiri::XML(file)
        end
      end
      @comments_doc = @comments_files.map do |item|
        File.open(item) do |file|
          Nokogiri::XML(file)
        end
      end
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
    @excelx_type = Hash.new
    @excelx_value = Hash.new
    @s_attribute = Hash.new # TODO: ggf. wieder entfernen nur lokal benoetigt
    @comment = Hash.new
    @comments_read = Hash.new
    @headers = first_row_peek(@default_sheet)
  end

  def method_missing(m,*args)
    # is method name a label name
    read_labels
    if @label.has_key?(m.to_s)
      sheet ||= @default_sheet
      read_cells(sheet) unless @cells_read[sheet]
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
    read_cells(sheet) unless @cells_read[sheet]
    row,col = normalize(row,col)
    if celltype(row,col,sheet) == :date
      yyyy,mm,dd = @cell[sheet][[row,col]].split('-')
      return Date.new(yyyy.to_i,mm.to_i,dd.to_i)
    elsif celltype(row,col,sheet) == :datetime
      date_part,time_part = @cell[sheet][[row,col]].split(' ')
      yyyy,mm,dd = date_part.split('-')
      hh,mi,ss = time_part.split(':')
      return DateTime.civil(yyyy.to_i,mm.to_i,dd.to_i,hh.to_i,mi.to_i,ss.to_i)
    end
    @cell[sheet][[row,col]]
  end

  # Returns the formula at (row,col).
  # Returns nil if there is no formula.
  # The method #formula? checks if there is a formula.
  def formula(row,col,sheet=nil)
    sheet ||= @default_sheet
    read_cells(sheet) unless @cells_read[sheet]
    row,col = normalize(row,col)
    if @formula[sheet][[row,col]] == nil
      return nil
    else
      return @formula[sheet][[row,col]]
    end
  end

  # true, if there is a formula
  def formula?(row,col,sheet=nil)
    sheet ||= @default_sheet
    read_cells(sheet) unless @cells_read[sheet]
    row,col = normalize(row,col)
    formula(row,col) != nil
  end

    # returns each formula in the selected sheet as an array of elements
  # [row, col, formula]
  def formulas(sheet=nil)
    sheet ||= @default_sheet
    read_cells(sheet) unless @cells_read[sheet]
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
      @bold == true
    end

    def italic?
      @italic == true
    end

    def underline?
      @underline == true
    end
  end

  # Given a cell, return the cell's style
  def font(row, col, sheet=nil)
    sheet ||= @default_sheet
    read_cells(sheet) unless @cells_read[sheet]
    row,col = normalize(row,col)
    s_attribute = @s_attribute[sheet][[row,col]]
    s_attribute ||= 0
    s_attribute = s_attribute.to_i
    @style_definitions[s_attribute]
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
    read_cells(sheet) unless @cells_read[sheet]
    row,col = normalize(row,col)
    if @formula[sheet][[row,col]]
      return :formula
    else
      @cell_type[sheet][[row,col]]
    end
  end

  # returns the internal type of an excel cell
  # * :numeric_or_formula
  # * :string
  # Note: this is only available within the Excelx class
  def excelx_type(row,col,sheet=nil)
    sheet ||= @default_sheet
    read_cells(sheet) unless @cells_read[sheet]
    row,col = normalize(row,col)
    return @excelx_type[sheet][[row,col]]
  end

  # returns the internal value of an excelx cell
  # Note: this is only available within the Excelx class
  def excelx_value(row,col,sheet=nil)
    sheet ||= @default_sheet
    read_cells(sheet) unless @cells_read[sheet]
    row,col = normalize(row,col)
    return @excelx_value[sheet][[row,col]]
  end

  # returns the internal format of an excel cell
  def excelx_format(row,col,sheet=nil)
    sheet ||= @default_sheet
    read_cells(sheet) unless @cells_read[sheet]
    row,col = normalize(row,col)
    s = @s_attribute[sheet][[row,col]]
    attribute2format(s).to_s
  end

  # returns an array of sheet names in the spreadsheet
  def sheets
    @workbook_doc.xpath("//xmlns:sheet").map do |sheet|
      sheet['name']
    end
  end

  # shows the internal representation of all cells
  # for debugging purposes
  def to_s(sheet=nil)
    sheet ||= @default_sheet
    read_cells(sheet) unless @cells_read[sheet]
    @cell[sheet].inspect
  end

  # returns the row,col values of the labelled cell
  # (nil,nil) if label is not defined
  def label(labelname)
    read_labels
    if @label.empty? || !@label.has_key?(labelname)
      return nil,nil,nil
    else
      return @label[labelname][1].to_i,
        Roo::GenericSpreadsheet.letter_to_number(@label[labelname][2]),
        @label[labelname][0]
    end
  end

  # Returns an array which all labels. Each element is an array with
  # [labelname, [row,col,sheetname]]
  def labels
    # sheet ||= @default_sheet
    # read_cells(sheet) unless @cells_read[sheet]
    read_labels
    @label.map do |label|
      [ label[0], # name
        [ label[1][1].to_i, # row
          Roo::GenericSpreadsheet.letter_to_number(label[1][2]), # column
          label[1][0], # sheet
        ] ]
    end
  end

  # returns the comment at (row/col)
  # nil if there is no comment
  def comment(row,col,sheet=nil)
    sheet ||= @default_sheet
    #read_cells(sheet) unless @cells_read[sheet]
    read_comments(sheet) unless @comments_read[sheet]
    row,col = normalize(row,col)
    return nil unless @comment[sheet]
    @comment[sheet][[row,col]]
  end

  # true, if there is a comment
  def comment?(row,col,sheet=nil)
    sheet ||= @default_sheet
    # read_cells(sheet) unless @cells_read[sheet]
    read_comments(sheet) unless @comments_read[sheet]
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

  # parse and return the first row without processing any of the rest of the sheet
  def first_row_peek(sheet=@default_sheet)
    validate_sheet!(sheet)
    read_row(sheet, @sheet_doc[sheets.index(sheet)].xpath("/xmlns:worksheet/xmlns:sheetData/xmlns:row").first, false)
  end

  # iterator for looping through rows without loading them all into memory
  def each_row(sheet=@default_sheet)
    validate_sheet!(sheet)
    @sheet_doc[sheets.index(sheet)].xpath("/xmlns:worksheet/xmlns:sheetData/xmlns:row").each_with_index { |r, i| yield read_row(sheet, r, false), i+1 }
  end

  private

  def formatted_value(value_type, v)
    case value_type
    when :float then v.to_f
    when :string then v
    when :date then (Date.new(1899,12,30)+v.to_i).strftime("%Y-%m-%d")
    when :datetime then (DateTime.new(1899,12,30)+v.to_f).strftime("%Y-%m-%d %H:%M:%S")
    when :percentage then v.to_f
    when :time then v.to_f*(24*60*60)
    else
      v
    end 
  end

  # helper function to set the internal representation of cells
  def set_cell_values(sheet,x,y,i,v,value_type,formula,
      excelx_type=nil,
      excelx_value=nil,
      s_attribute=nil)
    key = [y,x+i]
    @cell_type[sheet] ||= {}
    @cell_type[sheet][key] = value_type
    @formula[sheet] ||= {}
    @formula[sheet][key] = formula  if formula
    @cell[sheet] ||= {}
    @cell[sheet][key] = formatted_value(@cell_type[sheet][key])
    @excelx_type[sheet] ||= {}
    @excelx_type[sheet][key] = excelx_type
    @excelx_value[sheet] ||= {}
    @excelx_value[sheet][key] = excelx_value
    @s_attribute[sheet] ||= {}
    @s_attribute[sheet][key] = s_attribute
  end

  def format2type(format)
    format = format.to_s # weil von Typ Nokogiri::XML::Attr
    if FORMATS.has_key? format
      FORMATS[format]
    else
      #puts "FORMAT MISSING: #{format}"
      :float
    end
  end

  # search through the XML for the given cell to find and format values
  # if save_in_memory is true, these values will be saved using set_cell_values
  # if save_in_memory is false, these values will be returned as soon as they are discovered (for each_row feature)
  def read_cell(sheet=@default_sheet, c, save_in_memory)
    s_attribute = c['s'].to_i   # should be here
    # c: <c r="A5" s="2">
    # <v>22606</v>
    # </c>, format: , tmp_type: float
    y, x = Roo::GenericSpreadsheet.split_coordinate(c['r'])
    value_type =
      case c['t']
      when 's' then :shared
      when 'b' then :boolean
      when 'str' then :string
      when 'inlineStr' then :inlinestr
      else
        format = attribute2format(s_attribute)
        format2type(format)
      end
    formula = nil
    c.children.each do |cell|
      case cell.name
      when 'is'
        cell.children.each do |is|
          if is.name == 't'
            inlinestr_content = is.content
            value_type = :string
            v = inlinestr_content
            excelx_type = :string
            excelx_value = inlinestr_content #cell.content
            return [formatted_value(value_type,v),x] unless save_in_memory
            set_cell_values(sheet,x,y,0,v,value_type,formula,excelx_type,excelx_value,s_attribute)
          end
        end
      when 'f' then formula = cell.content
      when 'v'
        if [:time, :datetime].include?(value_type) && cell.content.to_f >= 1.0
          value_type = ((cell.content.to_f - cell.content.to_f.floor).abs > 0.000001) ? :datetime : :date
        end
        excelx_type = [:numeric_or_formula,format.to_s]
        excelx_value = cell.content
        v =
          case value_type
          when :shared
            value_type = :string
            excelx_type = :string
            @shared_table[cell.content.to_i]
          when :boolean then (cell.content.to_i == 1 ? 'TRUE' : 'FALSE')
          when :date, :time, :datetime then cell.content
          when :formula then cell.content.to_f #TODO: !!!!
          when :string
            excelx_type = :string
            cell.content
          else
            value_type = :float
            cell.content
          end
        return [formatted_value(value_type,v),x] unless save_in_memory
        set_cell_values(sheet,x,y,0,v,value_type,formula,excelx_type,excelx_value,s_attribute)
      end
    end
  end

  # read a single XML row from the selected sheet
  def read_row(sheet=@default_sheet, row_xml, save_in_memory)
    @headers.nil? ? ary = Array.new() : ary = Array.new(@headers.length) unless save_in_memory    
    row_xml.xpath("xmlns:c").each do |cell_xml|
      val = read_cell(sheet, cell_xml, save_in_memory)
      @headers.nil? ? ary << val[0] : ary[val[1]-1] = val[0] unless save_in_memory
    end
    return ary
  end

  # read all cells in the selected sheet
  def read_cells(sheet=@default_sheet)
    validate_sheet!(sheet)
    @sheet_doc[sheets.index(sheet)].xpath("/xmlns:worksheet/xmlns:sheetData/xmlns:row").each { |r| read_row(sheet, r, true) }
    @cells_read[sheet] = true
  end


  # Reads all comments from a sheet
  def read_comments(sheet=nil)
    sheet ||= @default_sheet
    validate_sheet!(sheet)
    n = self.sheets.index(sheet)
    return unless @comments_doc[n] #>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    @comments_doc[n].xpath("//xmlns:comments/xmlns:commentList/xmlns:comment").each do |comment|
      ref = comment.attributes['ref'].to_s
      row,col = Roo::GenericSpreadsheet.split_coordinate(ref)
      comment.xpath('./xmlns:text/xmlns:r/xmlns:t').each do |text|
        @comment[sheet] ||= {}
        @comment[sheet][[row,col]] = text.text
      end
    end
    @comments_read[sheet] = true
  end

  def read_labels
    @label ||= Hash[@workbook_doc.xpath("//xmlns:definedName").map do |defined_name|
	    # "Sheet1!$C$5"
      sheet, coordinates = defined_name.text.split('!$', 2)
      col,row = coordinates.split('$')
      [defined_name['name'], [sheet,row,col]]
    end]
  end

  # Extracts all needed files from the zip file
  def process_zipfile(tmpdir, zipfilename, zip, path='')
    @sheet_files = []
    Zip::ZipFile.open(zipfilename) {|zf|
      zf.entries.each {|entry|
        if entry.to_s.end_with?('workbook.xml')
          open(tmpdir+'/'+'roo_workbook.xml','wb') {|f|
            f << zip.read(entry)
          }
        end
        # if entry.to_s.end_with?('sharedStrings.xml')
	# at least one application creates this file with another (incorrect?)
	# casing. It doesn't hurt, if we ignore here the correct casing - there
	# won't be both names in the archive.
	# Changed the casing of all the following filenames.
        if entry.to_s.downcase.end_with?('sharedstrings.xml')
          open(tmpdir+'/'+'roo_sharedStrings.xml','wb') {|f|
            f << zip.read(entry)
          }
        end
        if entry.to_s.downcase.end_with?('styles.xml')
          open(tmpdir+'/'+'roo_styles.xml','wb') {|f|
            f << zip.read(entry)
          }
        end
        if entry.to_s.downcase =~ /sheet([0-9]+).xml$/
          nr = $1
          open(tmpdir+'/'+"roo_sheet#{nr}",'wb') {|f|
            f << zip.read(entry)
          }
          @sheet_files[nr.to_i-1] = tmpdir+'/'+"roo_sheet#{nr}"
        end
        if entry.to_s.downcase =~ /comments([0-9]+).xml$/
          nr = $1
          open(tmpdir+'/'+"roo_comments#{nr}",'wb') {|f|
            f << zip.read(entry)
          }
          @comments_files[nr.to_i-1] = tmpdir+'/'+"roo_comments#{nr}"
        end
      }
    }
    # return
  end

  # extract files from the zip file
  def extract_content(tmpdir, zipfilename)
    Zip::ZipFile.open(@filename) do |zip|
      process_zipfile(tmpdir, zipfilename,zip)
    end
  end

  # read the shared strings xml document
  def read_shared_strings(doc)
    doc.xpath("/xmlns:sst/xmlns:si").each do |si|
      shared_table_entry = ''
      si.children.each do |elem|
        if elem.name == 'r' and elem.children
          elem.children.each do |r_elem|
            if r_elem.name == 't'
              shared_table_entry << r_elem.content
            end
          end
        end
        if elem.name == 't'
          shared_table_entry = elem.content
        end
      end
      @shared_table << shared_table_entry
    end
  end

  # read the styles elements of an excelx document
  def read_styles(doc)
    @cellXfs = []

    @numFmts = Hash[doc.xpath("//xmlns:numFmt").map do |numFmt|
      [numFmt['numFmtId'], numFmt['formatCode']]
    end]
    fonts = doc.xpath("//xmlns:fonts/xmlns:font").map do |font_el|
      Font.new.tap do |font|
        font.bold = !font_el.xpath('./xmlns:b').empty?
        font.italic = !font_el.xpath('./xmlns:i').empty?
        font.underline = !font_el.xpath('./xmlns:u').empty?
      end
    end

    doc.xpath("//xmlns:cellXfs").each do |xfs|
      xfs.children.each do |xf|
        @cellXfs << xf['numFmtId']
        @style_definitions << fonts[xf['fontId'].to_i]
      end
    end
  end

  # convert internal excelx attribute to a format
  def attribute2format(s)
    id = @cellXfs[s.to_i]
    @numFmts[id] || STANDARD_FORMATS[id.to_i]
  end

end # class
