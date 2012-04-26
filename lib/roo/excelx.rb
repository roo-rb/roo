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
  @@nr = 0

  # initialization and opening of a spreadsheet file
  # values for packed: :zip
  def initialize(filename, packed=nil, file_warning = :error) #, create = false)
    super()
    @file_warning = file_warning
    file_type_check(filename,'.xlsx','an Excel-xlsx',packed)
    @tmpdir = Roo::GenericSpreadsheet.next_tmpdir
    @tmpdir = File.join(ENV['ROO_TMP'], @tmpdir) if ENV['ROO_TMP']
    unless File.exists?(@tmpdir)
      FileUtils::mkdir(@tmpdir)
    end
    filename = open_from_uri(filename) if filename[0,7] == "http://"
    filename = unzip(filename) if packed and packed == :zip
    @cells_read = Hash.new
    @filename = filename
    unless File.file?(@filename)
      FileUtils::rm_r(@tmpdir)
      raise IOError, "file #{@filename} does not exist"
    end
    @@nr += 1
    @file_nr = @@nr
    @comments_files = Array.new
    extract_content(@filename)
    @workbook_doc = File.open(File.join(@tmpdir, @file_nr.to_s+"_roo_workbook.xml")) do |file|
      Nokogiri::XML(file)
    end
    @shared_table = []
    if File.exist?(File.join(@tmpdir, @file_nr.to_s+'_roo_sharedStrings.xml'))
      @sharedstring_doc = File.open(File.join(@tmpdir, @file_nr.to_s+'_roo_sharedStrings.xml')) do |file|
        Nokogiri::XML(file)
      end
      read_shared_strings(@sharedstring_doc)
    end
    @styles_table = []
    @style_definitions = Array.new # TODO: ??? { |h,k| h[k] = {} }
    if File.exist?(File.join(@tmpdir, @file_nr.to_s+'_roo_styles.xml'))
      @styles_doc = File.open(File.join(@tmpdir, @file_nr.to_s+'_roo_styles.xml')) do |file|
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
    FileUtils::rm_r(@tmpdir)
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
  end

  def method_missing(m,*args)
    # is method name a label name
    read_labels
    if @label.has_key?(m.to_s)
      sheet = @default_sheet unless sheet
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
    sheet = @default_sheet unless sheet
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
    sheet = @default_sheet unless sheet
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
    sheet = @default_sheet unless sheet
    read_cells(sheet) unless @cells_read[sheet]
    row,col = normalize(row,col)
    formula(row,col) != nil
  end

    # returns each formula in the selected sheet as an array of elements
  # [row, col, formula]
  def formulas(sheet=nil)
    sheet = @default_sheet unless sheet
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
    sheet = @default_sheet unless sheet
    read_cells(sheet) unless @cells_read[sheet]
    row,col = normalize(row,col)
    s_attribute = @s_attribute[sheet][[row,col]]
    s_attribute ||= 0
    s_attribute = s_attribute.to_i
    @style_definitions[s_attribute]
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
    sheet = @default_sheet unless sheet
    read_cells(sheet) unless @cells_read[sheet]
    row,col = normalize(row,col)
    return @excelx_type[sheet][[row,col]]
  end

  # returns the internal value of an excelx cell
  # Note: this is only available within the Excelx class
  def excelx_value(row,col,sheet=nil)
    sheet = @default_sheet unless sheet
    read_cells(sheet) unless @cells_read[sheet]
    row,col = normalize(row,col)
    return @excelx_value[sheet][[row,col]]
  end

  # returns the internal format of an excel cell
  def excelx_format(row,col,sheet=nil)
    sheet = @default_sheet unless sheet
    read_cells(sheet) unless @cells_read[sheet]
    row,col = normalize(row,col)
    s = @s_attribute[sheet][[row,col]]
    result = attribute2format(s).to_s
    result
  end

  # returns an array of sheet names in the spreadsheet
  def sheets
    return_sheets = []
    @workbook_doc.xpath("//*[local-name()='sheet']").each do |sheet|
      return_sheets << sheet['name']
    end
    return_sheets
  end

  # shows the internal representation of all cells
  # for debugging purposes
  def to_s(sheet=nil)
    sheet = @default_sheet unless sheet
    read_cells(sheet) unless @cells_read[sheet]
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
        Roo::GenericSpreadsheet.letter_to_number(@label[labelname][2]),
        @label[labelname][0]
    else
      return nil,nil,nil
    end
  end

  # Returns an array which all labels. Each element is an array with
  # [labelname, [row,col,sheetname]]
  def labels
    # sheet = @default_sheet unless sheet
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
    sheet = @default_sheet unless sheet
    #read_cells(sheet) unless @cells_read[sheet]
    read_comments(sheet) unless @comments_read[sheet]
    row,col = normalize(row,col)
    return nil unless @comment[sheet]
    @comment[sheet][[row,col]]
  end

  # true, if there is a comment
  def comment?(row,col,sheet=nil)
    sheet = @default_sheet unless sheet
    # read_cells(sheet) unless @cells_read[sheet]
    read_comments(sheet) unless @comments_read[sheet]
    row,col = normalize(row,col)
    comment(row,col) != nil
  end

  # returns each comment in the selected sheet as an array of elements
  # [row, col, comment]
  def comments(sheet=nil)
    sheet = @default_sheet unless sheet
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

  # helper function to set the internal representation of cells
  def set_cell_values(sheet,x,y,i,v,value_type,formula,tr,str_v,
      excelx_type=nil,
      excelx_value=nil,
      s_attribute=nil)
    key = [y,x+i]
    @cell_type[sheet] = {} unless @cell_type[sheet]
    @cell_type[sheet][key] = value_type
    @formula[sheet] = {} unless @formula[sheet]
    @formula[sheet][key] = formula  if formula
    @cell[sheet]    = {} unless @cell[sheet]
    case @cell_type[sheet][key]
    when :float
      @cell[sheet][key] = v.to_f
    when :string
      @cell[sheet][key] = str_v
    when :date
      @cell[sheet][key] = (Date.new(1899,12,30)+v.to_i).strftime("%Y-%m-%d")
    when :datetime
      @cell[sheet][key] = (DateTime.new(1899,12,30)+v.to_f).strftime("%Y-%m-%d %H:%M:%S")
    when :percentage
      @cell[sheet][key] = v.to_f
    when :time
      @cell[sheet][key] = v.to_f*(24*60*60)
    else
      @cell[sheet][key] = v
    end
    @excelx_type[sheet] = {} unless @excelx_type[sheet]
    @excelx_type[sheet][key] = excelx_type
    @excelx_value[sheet] = {} unless @excelx_value[sheet]
    @excelx_value[sheet][key] = excelx_value
    @s_attribute[sheet] = {} unless @s_attribute[sheet]
    @s_attribute[sheet][key] = s_attribute
  end

  def format2type(format)
    format = format.to_s # weil von Typ Nokogiri::XML::Attr
    if FORMATS.has_key? format
      FORMATS[format]
    else
      :float
    end
  end

  # read all cells in the selected sheet
  def read_cells(sheet=nil)
    sheet = @default_sheet unless sheet
    sheet_found = false
    raise ArgumentError, "Error: sheet '#{sheet||'nil'}' not valid" if @default_sheet == nil and sheet==nil
    raise RangeError unless self.sheets.include? sheet
    n = self.sheets.index(sheet)
    @sheet_doc[n].xpath("//*[local-name()='c']").each do |c|
      s_attribute = c['s'].to_i   # should be here
      # c: <c r="A5" s="2">
      # <v>22606</v>
      # </c>, format: , tmp_type: float
      value_type =
        case c['t']
        when 's'
          :shared
        when 'b'
          :boolean
          # 2011-02-25 BEGIN
        when 'str'
          :string
          # 2011-02-25 END
          # 2011-09-15 BEGIN
        when 'inlineStr'
  	      :inlinestr
          # 2011-09-15 END
        else
          format = attribute2format(s_attribute)
          format2type(format)
        end
      formula = nil
      c.children.each do |cell|
	      # 2011-09-15 BEGIN
        if cell.name == 'is'
          cell.children.each do |is|
            if is.name == 't'
              inlinestr_content = is.content
              value_type = :string
              str_v = inlinestr_content
              excelx_type = :string
              y, x = Roo::GenericSpreadsheet.split_coordinate(c['r'])
              v = nil
              tr=nil #TODO: ???s
              excelx_value = inlinestr_content #cell.content
              set_cell_values(sheet,x,y,0,v,value_type,formula,tr,str_v,excelx_type,excelx_value,s_attribute)
            end
          end
        end
	      # 2011-09-15 END
        if cell.name == 'f'
          formula = cell.content
        end
        if cell.name == 'v'
          if value_type == :time or value_type == :datetime
            if cell.content.to_f >= 1.0
              if (cell.content.to_f - cell.content.to_f.floor).abs > 0.000001
                value_type = :datetime
              else
                value_type = :date
              end
            else
            end
          end
          excelx_type = [:numeric_or_formula,format.to_s]
          excelx_value = cell.content
          case value_type
          when :shared
            value_type = :string
            str_v = @shared_table[cell.content.to_i]
            excelx_type = :string
          when :boolean
            cell.content.to_i == 1 ? v = 'TRUE' : v = 'FALSE'
          when :date
            v = cell.content
          when :time
            v = cell.content
          when :datetime
            v = cell.content
          when :formula
            value_type = :formula
            v = cell.content.to_f #TODO: !!!!
            # 2011-02-25 BEGIN
          when :string
            str_v = cell.content
            excelx_type = :string
            # 2011-02-25 END
          else
            value_type = :float
            v = cell.content
          end
          y, x = Roo::GenericSpreadsheet.split_coordinate(c['r'])
          tr=nil #TODO: ???s
          set_cell_values(sheet,x,y,0,v,value_type,formula,tr,str_v,excelx_type,excelx_value,s_attribute)
        end
      end
    end
    sheet_found = true #TODO:
    if !sheet_found
      raise RangeError
    end
    @cells_read[sheet] = true
    # begin comments
=begin
Datei xl/comments1.xml
  <?xml version="1.0" encoding="UTF-8" standalone="yes" ?>
  <comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <authors>
      <author />
    </authors>
    <commentList>
      <comment ref="B4" authorId="0">
        <text>
          <r>
            <rPr>
              <sz val="10" />
              <rFont val="Arial" />
              <family val="2" />
            </rPr>
            <t>Kommentar fuer B4</t>
          </r>
        </text>
      </comment>
      <comment ref="B5" authorId="0">
        <text>
          <r>
            <rPr>
            <sz val="10" />
            <rFont val="Arial" />
            <family val="2" />
          </rPr>
          <t>Kommentar fuer B5</t>
        </r>
      </text>
    </comment>
  </commentList>
  </comments>
=end
=begin
    if @comments_doc[self.sheets.index(sheet)]
      read_comments(sheet)
    end
=end
    #end comments
  end

  # Reads all comments from a sheet
  def read_comments(sheet=nil)
    sheet = @default_sheet unless sheet
    #sheet_found = false
    raise ArgumentError, "Error: sheet '#{sheet||'nil'}' not valid" if @default_sheet == nil and sheet==nil
    raise RangeError unless self.sheets.include? sheet
    n = self.sheets.index(sheet)
    return unless @comments_doc[n] #>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    @comments_doc[n].xpath("//*[local-name()='comments']").each do |comment|
      comment.children.each do |cc|
        if cc.name == 'commentList'
          cc.children.each do |commentlist|
            if commentlist.name == 'comment'
              ref = commentlist.attributes['ref'].to_s
              row,col = Roo::GenericSpreadsheet.split_coordinate(ref)
              commentlist.children.each do |clc|
                if clc.name == 'text'
                  clc.children.each do |text|
                    if text.name == 'r'
                      text.children.each do |r|
                        if r.name == 't'
                          comment = r.text
                          @comment[sheet] = Hash.new unless @comment[sheet]
                          @comment[sheet][[row,col]] = comment
                        end
                      end
                    end
                  end
                end
              end
            end
          end
        end
      end
    end
    @comments_read[sheet] = true
  end

  def read_labels
    @label ||= Hash[@workbook_doc.xpath("//*[local-name()='definedName']").map do |defined_name|
	    # "Sheet1!$C$5"
      sheet, coordinates = defined_name.text.split('!$', 2)
      col,row = coordinates.split('$')
      [defined_name['name'], [sheet,row,col]]
    end]
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

  # Extracts all needed files from the zip file
  def process_zipfile(zipfilename, zip, path='')
    @sheet_files = []
    Zip::ZipFile.open(zipfilename) {|zf|
      zf.entries.each {|entry|
        if entry.to_s.end_with?('workbook.xml')
          open(@tmpdir+'/'+@file_nr.to_s+'_roo_workbook.xml','wb') {|f|
            f << zip.read(entry)
          }
        end
        # if entry.to_s.end_with?('sharedStrings.xml')
	# at least one application creates this file with another (incorrect?)
	# casing. It doesn't hurt, if we ignore here the correct casing - there
	# won't be both names in the archive.
	# Changed the casing of all the following filenames.
        if entry.to_s.downcase.end_with?('sharedstrings.xml')
          open(@tmpdir+'/'+@file_nr.to_s+'_roo_sharedStrings.xml','wb') {|f|
            f << zip.read(entry)
          }
        end
        if entry.to_s.downcase.end_with?('styles.xml')
          open(@tmpdir+'/'+@file_nr.to_s+'_roo_styles.xml','wb') {|f|
            f << zip.read(entry)
          }
        end
        if entry.to_s.downcase =~ /sheet([0-9]+).xml$/
          nr = $1
          open(@tmpdir+'/'+@file_nr.to_s+"_roo_sheet#{nr}",'wb') {|f|
            f << zip.read(entry)
          }
          @sheet_files[nr.to_i-1] = @tmpdir+'/'+@file_nr.to_s+"_roo_sheet#{nr}"
        end
        if entry.to_s.downcase =~ /comments([0-9]+).xml$/
          nr = $1
          open(@tmpdir+'/'+@file_nr.to_s+"_roo_comments#{nr}",'wb') {|f|
            f << zip.read(entry)
          }
          @comments_files[nr.to_i-1] = @tmpdir+'/'+@file_nr.to_s+"_roo_comments#{nr}"
        end
      }
    }
    # return
  end

  # extract files from the zip file
  def extract_content(zipfilename)
    Zip::ZipFile.open(@filename) do |zip|
      process_zipfile(zipfilename,zip)
    end
  end

  # sets the value of a cell
  def set_value(row,col,value,sheet=nil)
    sheet = @default_value unless sheet
    @cell[sheet][[row,col]] = value
  end

  # sets the type of a cell
  def set_type(row,col,type,sheet=nil)
    sheet = @default_value unless sheet
    @cell_type[sheet][[row,col]] = type
  end

  # read the shared strings xml document
  def read_shared_strings(doc)
    doc.xpath("//*[local-name()='si']").each do |si|
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
    @numFmts = []
    @cellXfs = []
    fonts = []

    doc.xpath("//*[local-name()='numFmt']").each do |numFmt|
      numFmtId = numFmt.attributes['numFmtId']
      formatCode = numFmt.attributes['formatCode']
      @numFmts << [numFmtId, formatCode]
    end
    doc.xpath("//*[local-name()='fonts']").each do |fonts_el|
      fonts_el.children.each do |font_el|
        if font_el == 'font'
          font = Excelx::Font.new
          font_el.each_element do |font_sub_el|
            case font_sub_el.name
            when 'b'
              font.bold = true
            when 'i'
              font.italic = true
            when 'u'
              font.underline = true
            end
          end
          fonts << font
        end
      end
    end

    doc.xpath("//*[local-name()='cellXfs']").each do |xfs|
      xfs.children.each do |xf|
        numFmtId = xf['numFmtId']
        @cellXfs << [numFmtId]
        fontId = xf['fontId'].to_i
        @style_definitions << fonts[fontId]
      end
    end
  end

  # convert internal excelx attribute to a format
  def attribute2format(s)
    result = nil
    @numFmts.each {|nf|
      # to_s weil das eine Nokogiri::XML::Attr und das
      # andere ein String ist
      if nf.first.to_s == @cellXfs[s.to_i].first
        result = nf[1]
        break
      end
    }
    unless result
      id = @cellXfs[s.to_i].first.to_i
      if STANDARD_FORMATS.has_key? id
        result = STANDARD_FORMATS[id]
      end
    end
    result
  end

end # class
