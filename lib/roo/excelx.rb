require 'date'
require 'nokogiri'
require 'roo/link'

class Roo::Excelx < Roo::Base
  module Format
    EXCEPTIONAL_FORMATS = {
      'h:mm am/pm' => :date,
      'h:mm:ss am/pm' => :date,
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

    def to_type(format)
      format = format.to_s.downcase
      if type = EXCEPTIONAL_FORMATS[format]
        type
      elsif format.include?('#')
        :float
      elsif !format.match(/d+(?![\]])/).nil? || format.include?('y')
        if format.include?('h') || format.include?('s')
          :datetime
        else
          :date
        end
      elsif format.include?('h') || format.include?('s')
        :time
      elsif format.include?('%')
        :percentage
      else
        :float
      end
    end

    module_function :to_type
  end

  class Sheet
    def initialize(name, rels_doc, sheet_doc, comments_doc)
      @name = name
      @rels_doc = rels_doc
      @sheet_doc = sheet_doc
      @comments_doc = comments_doc
    end

    def comment(key)
      comments[key]
    end

    def comments
      @comments ||=
        if @comments_doc
          Hash[@comments_doc.xpath("//comments/commentList/comment").map do |comment|
            [ref_to_key(comment), comment.at_xpath('./text/r/t').text ]
          end]
        else
          {}
        end
    end

    def hyperlink(key)
      hyperlinks[key]
    end

    private

    def ref_to_key(element)
      Roo::Base.split_coordinate(element.attributes['ref'].to_s)
    end

    def hyperlinks
      @hyperlinks ||=
        Hash[@sheet_doc.xpath("/worksheet/hyperlinks/hyperlink").map do |hyperlink|
          if hyperlink.attribute('id') && relationship = relationships[hyperlink.attribute('id').text]
            [ref_to_key(hyperlink), relationship.attribute('Target').text]
          end
        end.compact]
    end

    def relationships
      @relationships ||=
        if @rels_doc
          Hash[@rels_doc.xpath("/Relationships/Relationship").map do |rel|
            [rel.attribute('Id').text, rel]
          end]
        end
    end
  end

  # initialization and opening of a spreadsheet file
  # values for packed: :zip
  def initialize(filename, options = {})
    packed = options[:packed]
    file_warning = options[:file_warning] || :error

    file_type_check(filename,'.xlsx','an Excel-xlsx', file_warning, packed)

    @tmpdir = make_tmpdir(filename.split('/').last, options[:tmpdir_root])
    @filename = local_filename(filename, @tmpdir, packed)
    @comments_files = []
    @rels_files = []
    process_zipfile(@tmpdir, @filename)
    @sheet_doc = load_xmls(@sheet_files)
    @comments_doc = load_xmls(@comments_files)
    @rels_doc = load_xmls(@rels_files)

    super(filename, options)
    @formula = {}
    @excelx_type = {}
    @excelx_value = {}
    @style = {}
  end

  def method_missing(m,*args)
    # is method name a label name
    read_labels
    if @label.has_key?(m.to_s)
      sheet ||= default_sheet
      read_cells(sheet)
      row,col = label(m.to_s)
      cell(row,col)
    else
      # call super for methods like #a1
      super
    end
  end

  def sheet_for(sheet)
    sheet ||= default_sheet
    validate_sheet!(sheet)
    n = self.sheets.index(sheet)

    Sheet.new(sheet, @rels_doc[n], @sheet_doc[n], @comments_doc[n])
  end

  # Returns the content of a spreadsheet-cell.
  # (1,1) is the upper left corner.
  # (1,1), (1,'A'), ('A',1), ('a',1) all refers to the
  # cell at the first line and first row.
  def cell(row, col, sheet=nil)
    sheet ||= default_sheet
    sheet_object = sheet_for(sheet)
    read_cells(sheet)
    row,col = key = normalize(row,col)
    case celltype(row,col,sheet)
    when :date
      yyyy,mm,dd = @cell[sheet][key].split('-')
      Date.new(yyyy.to_i,mm.to_i,dd.to_i)
    when :datetime
      create_datetime_from(@cell[sheet][key])
    when :link
      Roo::Link.new(sheet_object.hyperlink(key), @cell[sheet][key].to_s)
    else
      @cell[sheet][key]
    end
  end

  # Returns the formula at (row,col).
  # Returns nil if there is no formula.
  # The method #formula? checks if there is a formula.
  def formula(row,col,sheet=nil)
    sheet ||= default_sheet
    read_cells(sheet)
    row,col = normalize(row,col)
    @formula[sheet][[row,col]] && @formula[sheet][[row,col]]
  end
  alias_method :formula?, :formula

    # returns each formula in the selected sheet as an array of elements
  # [row, col, formula]
  def formulas(sheet=nil)
    sheet ||= default_sheet
    read_cells(sheet)
    if @formula[sheet]
      @formula[sheet].map do |coord, formula|
        [coord[0], coord[1], formula]
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
    sheet ||= default_sheet
    read_cells(sheet)
    row,col = normalize(row,col)
    style_definitions[@style[sheet][[row,col]].to_i]
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
    sheet ||= default_sheet
    read_cells(sheet)
    sheet_object = sheet_for(sheet)

    key = normalize(row,col)
    if @formula[sheet][key]
      :formula
    elsif sheet_object.hyperlink(key)
      :link
    else
      @cell_type[sheet][key]
    end
  end

  # returns the internal type of an excel cell
  # * :numeric_or_formula
  # * :string
  # Note: this is only available within the Excelx class
  def excelx_type(row,col,sheet=nil)
    sheet ||= default_sheet
    read_cells(sheet)
    row,col = normalize(row,col)
    @excelx_type[sheet][[row,col]]
  end

  # returns the internal value of an excelx cell
  # Note: this is only available within the Excelx class
  def excelx_value(row,col,sheet=nil)
    sheet ||= default_sheet
    read_cells(sheet)
    row,col = normalize(row,col)
    @excelx_value[sheet][[row,col]]
  end

  # returns the internal format of an excel cell
  def excelx_format(row,col,sheet=nil)
    sheet ||= default_sheet
    read_cells(sheet)
    row,col = normalize(row,col)
    style_format(@style[sheet][[row,col]]).to_s
  end

  # returns an array of sheet names in the spreadsheet
  def sheets
    workbook_doc.xpath("//sheet").map do |sheet|
      sheet['name']
    end
  end

  # shows the internal representation of all cells
  # for debugging purposes
  def to_s(sheet=nil)
    sheet ||= default_sheet
    read_cells(sheet)
    @cell[sheet].inspect
  end

  # returns the row,col values of the labelled cell
  # (nil,nil) if label is not defined
  def label(labelname)
    read_labels
    if @label.empty? || !@label.has_key?(labelname)
      [nil,nil,nil]
    else
      [@label[labelname][1].to_i,
        self.class.letter_to_number(@label[labelname][2]),
        @label[labelname][0]]
    end
  end

  # Returns an array which all labels. Each element is an array with
  # [labelname, [row,col,sheetname]]
  def labels
    # sheet ||= default_sheet
    # read_cells(sheet)
    read_labels
    @label.map do |label|
      [ label[0], # name
        [ label[1][1].to_i, # row
          self.class.letter_to_number(label[1][2]), # column
          label[1][0], # sheet
        ] ]
    end
  end

  def hyperlink?(row,col,sheet=nil)
    hyperlink(row, col, sheet) != nil
  end

  # returns the hyperlink at (row/col)
  # nil if there is no hyperlink
  def hyperlink(row,col,sheet=nil)
    key = normalize(row,col)
    sheet_for(sheet).hyperlink(key)
  end

  # returns the comment at (row/col)
  # nil if there is no comment
  def comment(row,col,sheet=nil)
    key = normalize(row,col)
    sheet_for(sheet).comment(key)
  end

  # true, if there is a comment
  def comment?(row,col,sheet=nil)
    comment(row,col,sheet) != nil
  end

  def comments(sheet=nil)
    sheet_for(sheet).comments.map do |(x, y), comment|
      [x, y, comment]
    end
  end

  private

  def workbook_doc
    @workbook_doc ||= load_xml(File.join(@tmpdir, "roo_workbook.xml"))
  end

  def load_xml(path)
    super.remove_namespaces!
  end

  def load_xmls(paths)
    paths.compact.map do |item|
      load_xml(item).remove_namespaces!
    end
  end

  # helper function to set the internal representation of cells
  def set_cell_values(sheet,x,y,i,v,value_type,formula,
      excelx_type=nil,
      excelx_value=nil,
      style=nil)
    key = [y,x+i]
    @cell_type[sheet] ||= {}
    @cell_type[sheet][key] = value_type
    @formula[sheet] ||= {}
    @formula[sheet][key] = formula if formula
    @cell[sheet] ||= {}
    @cell[sheet][key] =
      case @cell_type[sheet][key]
      when :float
        v.to_f
      when :string
        v
      when :date
        (base_date+v.to_i).strftime("%Y-%m-%d")
      when :datetime
        (base_date+v.to_f.round(6)).strftime("%Y-%m-%d %H:%M:%S.%N")
      when :percentage
        v.to_f
      when :time
        v.to_f*(24*60*60)
      else
        v
      end

    @excelx_type[sheet] ||= {}
    @excelx_type[sheet][key] = excelx_type
    @excelx_value[sheet] ||= {}
    @excelx_value[sheet][key] = excelx_value
    @style[sheet] ||= {}
    @style[sheet][key] = style
  end

  def read_cell_from_xml(sheet, cell_xml)
    style = cell_xml['s'].to_i   # should be here
    # c: <c r="A5" s="2">
    # <v>22606</v>
    # </c>, format: , tmp_type: float
    value_type =
      case cell_xml['t']
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
        format = style_format(style)
        Format.to_type(format)
      end
    formula = nil
    cell_xml.children.each do |cell|
      case cell.name
        when 'is'
          cell.children.each do |is|
            if is.name == 't'
              inlinestr_content = is.content
              value_type = :string
              v = inlinestr_content
              excelx_type = :string
              y, x = self.class.split_coordinate(cell_xml['r'])
              excelx_value = inlinestr_content #cell.content
              set_cell_values(sheet,x,y,0,v,value_type,formula,excelx_type,excelx_value,style)
            end
          end
        when 'f'
          formula = cell.content
        when 'v'
          if [:time, :datetime].include?(value_type) && cell.content.to_f >= 1.0
            value_type =
                if (cell.content.to_f - cell.content.to_f.floor).abs > 0.000001
                  :datetime
                else
                  :date
                end
          end
          excelx_type = [:numeric_or_formula,format.to_s]
          excelx_value = cell.content
          v =
            case value_type
            when :shared
              value_type = :string
              excelx_type = :string
              shared_strings[cell.content.to_i]
            when :boolean
              (cell.content.to_i == 1 ? 'TRUE' : 'FALSE')
            when :date
              cell.content
            when :time
              cell.content
            when :datetime
              cell.content
            when :formula
              cell.content.to_f #TODO: !!!!
            when :string
              excelx_type = :string
              cell.content
            else
              value_type = :float
              cell.content
            end
          y, x = self.class.split_coordinate(cell_xml['r'])
          set_cell_values(sheet,x,y,0,v,value_type,formula,excelx_type,excelx_value,style)
      end
    end
  end

  # read all cells in the selected sheet
  def read_cells(sheet=nil)
    sheet ||= default_sheet
    validate_sheet!(sheet)
    return if @cells_read[sheet]

    @sheet_doc[sheets.index(sheet)].xpath("/worksheet/sheetData/row/c").each do |c|
      read_cell_from_xml(sheet, c)
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

  def read_labels
    @label ||= Hash[workbook_doc.xpath("//definedName").map do |defined_name|
      # "Sheet1!$C$5"
      sheet, coordinates = defined_name.text.split('!$', 2)
      col,row = coordinates.split('$')
      [defined_name['name'], [sheet,row,col]]
    end]
  end

  # Extracts all needed files from the zip file
  def process_zipfile(tmpdir, zipfilename)
    @sheet_files = []
    Roo::ZipFile.open(zipfilename) do |zipfile|
      zipfile.entries.each do |entry|
        entry_name = entry.to_s.downcase

        path =
          case entry_name
          when /workbook.xml$/
            "#{tmpdir}/roo_workbook.xml"
          when /sharedstrings.xml$/
            "#{tmpdir}/roo_sharedStrings.xml"
          when /styles.xml$/
            "#{tmpdir}/roo_styles.xml"
          when /sheet.xml$/
            path = "#{tmpdir}/roo_sheet"
            @sheet_files.unshift path
            path
          when /sheet([0-9]+).xml$/
            # Numbers 3.1 exports first sheet without sheet number. Such sheets
            # are always added to the beginning of the array which, naturally,
            # causes other sheets to be pushed to the next index which could
            # lead to sheet references getting overwritten, so we need to
            # handle that case specifically.
            nr = $1
            sheet_files_index = nr.to_i - 1
            sheet_files_index += 1 if @sheet_files[sheet_files_index]
            @sheet_files[sheet_files_index] = "#{tmpdir}/roo_sheet#{nr.to_i}"
          when /comments([0-9]+).xml$/
            nr = $1
            @comments_files[nr.to_i-1] = "#{tmpdir}/roo_comments#{nr}"
          when /sheet([0-9]+).xml.rels$/
            nr = $1
            @rels_files[nr.to_i-1] = "#{tmpdir}/roo_rels#{nr}"
          end
        if path
          File.write(path, zipfile.read(entry), mode: 'wb')
        end
      end
    end
  end

  def shared_strings
    @shared_strings ||=
      if File.exist?(shared_strings_path)
        # read the shared strings xml document
        xml = load_xml(shared_strings_path)
        xml.xpath("/sst/si").map do |si|
          shared_string = ''
          si.children.each do |elem|
            case elem.name
            when 'r'
              elem.children.each do |r_elem|
                if r_elem.name == 't'
                  shared_string << r_elem.content
                end
              end
            when 't'
              shared_string = elem.content
            end
          end
          shared_string
        end
      else
        []
      end
  end

  def shared_strings_path
    @shared_strings_path ||= File.join(@tmpdir, 'roo_sharedStrings.xml')
  end

  ##### STYLES
  def style_definitions
    @style_definitions ||= styles_doc.xpath("//cellXfs").flat_map do |xfs|
      xfs.children.map do |xf|
        fonts[xf['fontId'].to_i]
      end
    end
  end

  def num_fmt_ids
    @num_fmt_ids ||= styles_doc.xpath("//cellXfs").flat_map do |xfs|
      xfs.children.map do |xf|
        xf['numFmtId']
      end
    end
  end

  def num_fmts
    @num_fmts ||= Hash[styles_doc.xpath("//numFmt").map do |num_fmt|
      [num_fmt['numFmtId'], num_fmt['formatCode']]
    end]
  end

  def fonts
   @fonts ||= styles_doc.xpath("//fonts/font").map do |font_el|
      Font.new.tap do |font|
        font.bold = !font_el.xpath('./b').empty?
        font.italic = !font_el.xpath('./i').empty?
        font.underline = !font_el.xpath('./u').empty?
      end
    end
  end

  def styles_doc
    @styles_doc ||=
      if File.exist?(File.join(@tmpdir, 'roo_styles.xml'))
        load_xml(File.join(@tmpdir, 'roo_styles.xml'))
      end
  end

  # convert internal excelx attribute to a format
  def style_format(style)
    id = num_fmt_ids[style.to_i]
    num_fmts[id] || Format::STANDARD_FORMATS[id.to_i]
  end
  ###### END STYLES

  def base_date
    @base_date ||=
      begin
        # Default to 1900 (minus one day due to excel quirk) but use 1904 if
        # it's set in the Workbook's workbookPr
        # http://msdn.microsoft.com/en-us/library/ff530155(v=office.12).aspx
        workbook_doc.css("workbookPr[date1904]").each do |workbookPr|
          if workbookPr["date1904"] =~ /true|1/i
            return Date.new(1904,01,01)
          end
        end
        Date.new(1899,12,30)
      end
  end

  def create_datetime_from(datetime_string)
    date_part,time_part = round_time_from(datetime_string).split(' ')
    yyyy,mm,dd = date_part.split('-')
    hh,mi,ss = time_part.split(':')
    DateTime.civil(yyyy.to_i,mm.to_i,dd.to_i,hh.to_i,mi.to_i,ss.to_i)
  end

  def round_time_from(datetime_string)
    date_part,time_part = datetime_string.split(' ')
    yyyy,mm,dd = date_part.split('-')
    hh,mi,ss = time_part.split(':')
    Time.new(yyyy.to_i, mm.to_i, dd.to_i, hh.to_i, mi.to_i, ss.to_r).round(0).strftime("%Y-%m-%d %H:%M:%S")
  end
end
