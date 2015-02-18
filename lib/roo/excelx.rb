require 'date'
require 'nokogiri'
require 'roo/link'
require 'roo/utils'
require 'zip/filesystem'

class Roo::Excelx < Roo::Base
  autoload :Workbook, 'roo/excelx/workbook'
  autoload :SharedStrings, 'roo/excelx/shared_strings'
  autoload :Styles, 'roo/excelx/styles'

  autoload :Relationships, 'roo/excelx/relationships'
  autoload :Comments, 'roo/excelx/comments'
  autoload :SheetDoc, 'roo/excelx/sheet_doc'

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
      if format == 'general'
        :string
      elsif type = EXCEPTIONAL_FORMATS[format]
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

  class Cell
    attr_reader :type, :formula, :value, :excelx_type, :excelx_value, :style, :hyperlink, :coordinate

    def initialize(value, type, formula, excelx_type, excelx_value, style, hyperlink, base_date, coordinate)
      @type = type
      @formula = formula
      @base_date = base_date if [:date, :datetime].include?(@type)
      @excelx_type = excelx_type
      @excelx_value = excelx_value
      @style = style
      @value = type_cast_value(value)
      @value = Roo::Link.new(hyperlink, @value.to_s) if hyperlink
      @coordinate = coordinate
    end

    def type
      if @formula
        :formula
      elsif @value.is_a?(Roo::Link)
        :link
      else
        @type
      end
    end

    class Coordinate
      attr_accessor :row, :column

      def initialize(row, column)
        @row, @column = row, column
      end
    end

    private

    def type_cast_value(value)
      case @type
      when :float, :percentage
        value.to_f
      when :date
        yyyy,mm,dd = (@base_date+value.to_i).strftime("%Y-%m-%d").split('-')
        Date.new(yyyy.to_i,mm.to_i,dd.to_i)
      when :datetime
        create_datetime_from((@base_date+value.to_f.round(6)).strftime("%Y-%m-%d %H:%M:%S.%N"))
      when :time
        value.to_f*(24*60*60)
      when :string
        value
      else
        value
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

  class Sheet
    def initialize(name, rels_path, sheet_path, comments_path, styles, shared_strings, workbook)
      @name = name
      @rels = Relationships.new(rels_path)
      @comments = Comments.new(comments_path)
      @styles = styles
      @sheet = SheetDoc.new(sheet_path, @rels, @styles, shared_strings, workbook)
    end

    def cells
      @cells ||= @sheet.cells(@rels)
    end

    def present_cells
      @present_cells ||= cells.select {|key, cell| cell && cell.value }
    end

    # Yield each row as array of Excelx::Cell objects
    # accepts options max_rows (int) (offset by 1 for header)
    # and pad_cells (boolean)
    def each_row(options = {}, &block)
      row_count = 0
      @sheet.each_row_streaming do |row|
        break if options[:max_rows] && row_count == options[:max_rows] + 1
        block.call(cells_for_row_element(row, options)) if block_given?
        row_count += 1
      end
    end

    def row(row_number)
      first_column.upto(last_column).map do |col|
        cells[[row_number,col]]
      end.map {|cell| cell && cell.value }
    end

    def column(col_number)
      first_row.upto(last_row).map do |row|
        cells[[row,col_number]]
      end.map {|cell| cell && cell.value }
    end

    # returns the number of the first non-empty row
    def first_row
      @first_row ||= present_cells.keys.map {|row, col| row }.min
    end

    def last_row
      @last_row ||= present_cells.keys.map {|row, col| row }.max
    end

    # returns the number of the first non-empty column
    def first_column(sheet=nil)
      @first_column ||= present_cells.keys.map {|row, col| col }.min
    end

    # returns the number of the last non-empty column
    def last_column(sheet=nil)
      @last_column ||= present_cells.keys.map {|row, col| col }.max
    end

    def excelx_format(key)
      @styles.style_format(cells[key].style).to_s
    end

    def hyperlinks
      @hyperlinks ||= @sheet.hyperlinks(@rels)
    end

    def comments
      @comments.comments
    end

    def dimensions
      @sheet.dimensions
    end

    private

    # Take an xml row and return an array of Excelx::Cell objects
    # optionally pad array to header width(assumed 1st row).
    # takes option pad_cells (boolean) defaults false
    def cells_for_row_element(row_element, options = {})
      return [] unless row_element
      cell_col = 0
      cells = []
      @sheet.each_cell(row_element) do |cell|
        cells.concat(pad_cells(cell, cell_col)) if options[:pad_cells]
        cells << cell
        cell_col = cell.coordinate.column
      end
      cells
    end

    def pad_cells(cell, last_column)
      pad = []
      (cell.coordinate.column - 1 - last_column).times { pad << nil }
      pad
    end
  end

  ExceedsMaxError = Class.new(StandardError)

  # initialization and opening of a spreadsheet file
  # values for packed: :zip
  # optional cell_max (int) parameter for early aborting attempts to parse
  # enormous documents.
  def initialize(filename, options = {})
    packed = options[:packed]
    file_warning = options.fetch(:file_warning, :error)
    cell_max = options.delete(:cell_max)

    file_type_check(filename,'.xlsx','an Excel-xlsx', file_warning, packed)

    @tmpdir = make_tmpdir(filename.split('/').last, options[:tmpdir_root])
    @filename = local_filename(filename, @tmpdir, packed)
    @comments_files = []
    @rels_files = []
    process_zipfile(@tmpdir, @filename)

    @sheet_names = workbook.sheets.map { |sheet| sheet['name'] }
    @sheets = []
    @sheets_by_name = Hash[@sheet_names.map.with_index do |sheet_name, n|
      @sheets[n] = Sheet.new(sheet_name, @rels_files[n], @sheet_files[n], @comments_files[n], styles, shared_strings, workbook)
      [sheet_name, @sheets[n]]
    end]

    if cell_max
      cell_count = ::Roo::Utils.num_cells_in_range(sheet_for(options.delete(:sheet)).dimensions)
      raise ExceedsMaxError.new("Excel file exceeds cell maximum: #{cell_count} > #{cell_max}") if cell_count > cell_max
    end

    super
  end

  def method_missing(method,*args)
    if label = workbook.defined_names[method.to_s]
      sheet_for(label.sheet).cells[label.key].value
    else
      # call super for methods like #a1
      super
    end
  end

  def sheets
    @sheet_names
  end

  def sheet_for(sheet)
    sheet ||= default_sheet
    validate_sheet!(sheet)
    @sheets_by_name[sheet]
  end

  # Returns the content of a spreadsheet-cell.
  # (1,1) is the upper left corner.
  # (1,1), (1,'A'), ('A',1), ('a',1) all refers to the
  # cell at the first line and first row.
  def cell(row, col, sheet=nil)
    key = normalize(row,col)
    cell = sheet_for(sheet).cells[key]
    cell.value if cell
  end

  def row(rownumber,sheet=nil)
    sheet_for(sheet).row(rownumber)
  end

  # returns all values in this column as an array
  # column numbers are 1,2,3,... like in the spreadsheet
  def column(column_number,sheet=nil)
    if column_number.is_a?(::String)
      column_number = ::Roo::Utils.letter_to_number(column_number)
    end
    sheet_for(sheet).column(column_number)
  end

  # returns the number of the first non-empty row
  def first_row(sheet=nil)
    sheet_for(sheet).first_row
  end

  # returns the number of the last non-empty row
  def last_row(sheet=nil)
    sheet_for(sheet).last_row
  end

  # returns the number of the first non-empty column
  def first_column(sheet=nil)
    sheet_for(sheet).first_column
  end

  # returns the number of the last non-empty column
  def last_column(sheet=nil)
    sheet_for(sheet).last_column
  end

  # set a cell to a certain value
  # (this will not be saved back to the spreadsheet file!)
  def set(row,col,value, sheet = nil) #:nodoc:
    key = normalize(row,col)
    cell_type = cell_type_by_value(value)
    sheet_for(sheet).cells[key] = Cell.new(value, cell_type, nil, cell_type, value, nil, nil, nil, Cell::Coordinate.new(row, col))
  end


  # Returns the formula at (row,col).
  # Returns nil if there is no formula.
  # The method #formula? checks if there is a formula.
  def formula(row,col,sheet=nil)
    key = normalize(row,col)
    sheet_for(sheet).cells[key].formula
  end

  # Predicate methods really should return a boolean
  # value. Hopefully no one was relying on the fact that this
  # previously returned either nil/formula
  def formula?(*args)
    !!formula(*args)
  end

  # returns each formula in the selected sheet as an array of tuples in following format
  # [[row, col, formula], [row, col, formula],...]
  def formulas(sheet=nil)
    sheet_for(sheet).cells.select {|_, cell| cell.formula }.map do |(x, y), cell|
      [x, y, cell.formula]
    end
  end

  # Given a cell, return the cell's style
  def font(row, col, sheet=nil)
    key = normalize(row,col)
    styles.definitions[sheet_for(sheet).cells[key].style]
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
    key = normalize(row, col)
    sheet_for(sheet).cells[key].type
  end

  # returns the internal type of an excel cell
  # * :numeric_or_formula
  # * :string
  # Note: this is only available within the Excelx class
  def excelx_type(row,col,sheet=nil)
    key = normalize(row,col)
    sheet_for(sheet).cells[key].excelx_type
  end

  # returns the internal value of an excelx cell
  # Note: this is only available within the Excelx class
  def excelx_value(row,col,sheet=nil)
    key = normalize(row,col)
    sheet_for(sheet).cells[key].excelx_value
  end

  # returns the internal format of an excel cell
  def excelx_format(row,col,sheet=nil)
    key = normalize(row,col)
    sheet_for(sheet).excelx_format(key)
  end

  def empty?(row,col,sheet=nil)
    sheet = sheet_for(sheet)
    key = normalize(row,col)
    cell = sheet.cells[key]
    !cell || !cell.value || (cell.type == :string && cell.value.empty?) \
      || (row < sheet.first_row || row > sheet.last_row || col < sheet.first_column || col > sheet.last_column)
  end

  # shows the internal representation of all cells
  # for debugging purposes
  def to_s(sheet=nil)
    sheet_for(sheet).cells.inspect
  end

  # returns the row,col values of the labelled cell
  # (nil,nil) if label is not defined
  def label(name)
    labels = workbook.defined_names
    if labels.empty? || !labels.key?(name)
      [nil,nil,nil]
    else
      [labels[name].row,
        labels[name].col,
        labels[name].sheet]
    end
  end

  # Returns an array which all labels. Each element is an array with
  # [labelname, [row,col,sheetname]]
  def labels
    @labels ||= workbook.defined_names.map do |name, label|
      [ name,
        [ label.row,
          label.col,
          label.sheet,
        ] ]
    end
  end

  def hyperlink?(row,col,sheet=nil)
    !!hyperlink(row, col, sheet)
  end

  # returns the hyperlink at (row/col)
  # nil if there is no hyperlink
  def hyperlink(row,col,sheet=nil)
    key = normalize(row,col)
    sheet_for(sheet).hyperlinks[key]
  end

  # returns the comment at (row/col)
  # nil if there is no comment
  def comment(row,col,sheet=nil)
    key = normalize(row,col)
    sheet_for(sheet).comments[key]
  end

  # true, if there is a comment
  def comment?(row,col,sheet=nil)
    !!comment(row,col,sheet)
  end

  def comments(sheet=nil)
    sheet_for(sheet).comments.map do |(x, y), comment|
      [x, y, comment]
    end
  end

  # Yield an array of Excelx::Cell
  # Takes options for sheet, pad_cells, and max_rows
  def each_row_streaming(options={})
    sheet_for(options.delete(:sheet)).each_row(options) { |row| yield row }
  end

  private

  # Extracts all needed files from the zip file
  def process_zipfile(tmpdir, zipfilename)
    @sheet_files = []
    Zip::File.foreach(zipfilename) do |entry|
      path =
        case entry.name.downcase
        when /workbook.xml$/
          "#{tmpdir}/roo_workbook.xml"
        when /sharedstrings.xml$/
          "#{tmpdir}/roo_sharedStrings.xml"
        when /styles.xml$/
          "#{tmpdir}/roo_styles.xml"
        when /sheet.xml$/
          path = "#{tmpdir}/roo_sheet"
          @sheet_files.unshift(path)
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
        entry.extract(path)
      end
    end
  end

  def styles
    @styles ||= Styles.new(File.join(@tmpdir, 'roo_styles.xml'))
  end

  def shared_strings
    @shared_strings ||= SharedStrings.new(File.join(@tmpdir, 'roo_sharedStrings.xml'))
  end

  def workbook
    @workbook ||= Workbook.new(File.join(@tmpdir, "roo_workbook.xml"))
  end
end
