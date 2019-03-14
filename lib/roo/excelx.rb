require 'nokogiri'
require 'zip/filesystem'
require 'roo/link'
require 'roo/tempdir'
require 'roo/utils'
require 'forwardable'
require 'set'

module Roo
  class Excelx < Roo::Base
    extend Roo::Tempdir
    extend Forwardable

    ERROR_VALUES = %w(#N/A #REF! #NAME? #DIV/0! #NULL! #VALUE! #NUM!).to_set

    require 'roo/excelx/shared'
    require 'roo/excelx/workbook'
    require 'roo/excelx/shared_strings'
    require 'roo/excelx/styles'
    require 'roo/excelx/cell'
    require 'roo/excelx/sheet'
    require 'roo/excelx/relationships'
    require 'roo/excelx/comments'
    require 'roo/excelx/sheet_doc'
    require 'roo/excelx/coordinate'
    require 'roo/excelx/format'
    require 'roo/excelx/images'

    delegate [:styles, :workbook, :shared_strings, :rels_files, :sheet_files, :comments_files, :image_rels, :image_files] => :@shared
    ExceedsMaxError = Class.new(StandardError)

    # initialization and opening of a spreadsheet file
    # values for packed: :zip
    # optional cell_max (int) parameter for early aborting attempts to parse
    # enormous documents.
    def initialize(filename_or_stream, options = {})
      packed = options[:packed]
      file_warning = options.fetch(:file_warning, :error)
      cell_max = options.delete(:cell_max)
      sheet_options = {}
      sheet_options[:expand_merged_ranges] = (options[:expand_merged_ranges] || false)
      sheet_options[:no_hyperlinks] = (options[:no_hyperlinks] || false)
      sheet_options[:empty_cell] = (options[:empty_cell] || false)
      shared_options = {}

      shared_options[:disable_html_wrapper] = (options[:disable_html_wrapper] || false)
      unless is_stream?(filename_or_stream)
        file_type_check(filename_or_stream, %w[.xlsx .xlsm], 'an Excel 2007', file_warning, packed)
        basename = find_basename(filename_or_stream)
      end

      # NOTE: Create temp directory and allow Ruby to cleanup the temp directory
      #       when the object is garbage collected. Initially, the finalizer was
      #       created in the Roo::Tempdir module, but that led to a segfault
      #       when testing in Ruby 2.4.0.
      @tmpdir = self.class.make_tempdir(self, basename, options[:tmpdir_root])
      ObjectSpace.define_finalizer(self, self.class.finalize(object_id))

      @shared = Shared.new(@tmpdir, shared_options)
      @filename = local_filename(filename_or_stream, @tmpdir, packed)
      process_zipfile(@filename || filename_or_stream)

      @sheet_names = workbook.sheets.map do |sheet|
        unless options[:only_visible_sheets] && sheet['state'] == 'hidden'
          sheet['name']
        end
      end.compact
      @sheets = []
      @sheets_by_name = {}
      @sheet_names.each_with_index do |sheet_name, n|
        @sheets_by_name[sheet_name] = @sheets[n] = Sheet.new(sheet_name, @shared, n, sheet_options)
      end

      if cell_max
        cell_count = ::Roo::Utils.num_cells_in_range(sheet_for(options.delete(:sheet)).dimensions)
        raise ExceedsMaxError.new("Excel file exceeds cell maximum: #{cell_count} > #{cell_max}") if cell_count > cell_max
      end

      super
    rescue
      self.class.finalize_tempdirs(object_id)
      raise
    end

    def method_missing(method, *args)
      if (label = workbook.defined_names[method.to_s])
        safe_send(sheet_for(label.sheet).cells[label.key], :value)
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
      @sheets_by_name[sheet] || @sheets[sheet]
    end

    def images(sheet = nil)
      images_names = sheet_for(sheet).images.map(&:last)
      images_names.map { |iname| image_files.find { |ifile| ifile[iname] } }
    end

    # Returns the content of a spreadsheet-cell.
    # (1,1) is the upper left corner.
    # (1,1), (1,'A'), ('A',1), ('a',1) all refers to the
    # cell at the first line and first row.
    def cell(row, col, sheet = nil)
      key = normalize(row, col)
      safe_send(sheet_for(sheet).cells[key], :value)
    end

    def row(rownumber, sheet = nil)
      sheet_for(sheet).row(rownumber)
    end

    # returns all values in this column as an array
    # column numbers are 1,2,3,... like in the spreadsheet
    def column(column_number, sheet = nil)
      if column_number.is_a?(::String)
        column_number = ::Roo::Utils.letter_to_number(column_number)
      end
      sheet_for(sheet).column(column_number)
    end

    # returns the number of the first non-empty row
    def first_row(sheet = nil)
      sheet_for(sheet).first_row
    end

    # returns the number of the last non-empty row
    def last_row(sheet = nil)
      sheet_for(sheet).last_row
    end

    # returns the number of the first non-empty column
    def first_column(sheet = nil)
      sheet_for(sheet).first_column
    end

    # returns the number of the last non-empty column
    def last_column(sheet = nil)
      sheet_for(sheet).last_column
    end

    # set a cell to a certain value
    # (this will not be saved back to the spreadsheet file!)
    def set(row, col, value, sheet = nil) #:nodoc:
      key = normalize(row, col)
      cell_type = cell_type_by_value(value)
      sheet_for(sheet).cells[key] = Cell.new(value, cell_type, nil, cell_type, value, nil, nil, nil, Coordinate.new(row, col))
    end

    # Returns the formula at (row,col).
    # Returns nil if there is no formula.
    # The method #formula? checks if there is a formula.
    def formula(row, col, sheet = nil)
      key = normalize(row, col)
      safe_send(sheet_for(sheet).cells[key], :formula)
    end

    # Predicate methods really should return a boolean
    # value. Hopefully no one was relying on the fact that this
    # previously returned either nil/formula
    def formula?(*args)
      !!formula(*args)
    end

    # returns each formula in the selected sheet as an array of tuples in following format
    # [[row, col, formula], [row, col, formula],...]
    def formulas(sheet = nil)
      sheet_for(sheet).cells.select { |_, cell| cell.formula }.map do |(x, y), cell|
        [x, y, cell.formula]
      end
    end

    # Given a cell, return the cell's style
    def font(row, col, sheet = nil)
      key = normalize(row, col)
      definition_index = safe_send(sheet_for(sheet).cells[key], :style)
      styles.definitions[definition_index] if definition_index
    end

    # returns the type of a cell:
    # * :float
    # * :string,
    # * :date
    # * :percentage
    # * :formula
    # * :time
    # * :datetime
    def celltype(row, col, sheet = nil)
      key = normalize(row, col)
      safe_send(sheet_for(sheet).cells[key], :type)
    end

    # returns the internal type of an excel cell
    # * :numeric_or_formula
    # * :string
    # Note: this is only available within the Excelx class
    def excelx_type(row, col, sheet = nil)
      key = normalize(row, col)
      safe_send(sheet_for(sheet).cells[key], :cell_type)
    end

    # returns the internal value of an excelx cell
    # Note: this is only available within the Excelx class
    def excelx_value(row, col, sheet = nil)
      key = normalize(row, col)
      safe_send(sheet_for(sheet).cells[key], :cell_value)
    end

    # returns the internal value of an excelx cell
    # Note: this is only available within the Excelx class
    def formatted_value(row, col, sheet = nil)
      key = normalize(row, col)
      safe_send(sheet_for(sheet).cells[key], :formatted_value)
    end

    # returns the internal format of an excel cell
    def excelx_format(row, col, sheet = nil)
      key = normalize(row, col)
      sheet_for(sheet).excelx_format(key)
    end

    def empty?(row, col, sheet = nil)
      sheet = sheet_for(sheet)
      key = normalize(row, col)
      cell = sheet.cells[key]
      !cell || cell.empty? ||
        (row < sheet.first_row || row > sheet.last_row || col < sheet.first_column || col > sheet.last_column)
    end

    # shows the internal representation of all cells
    # for debugging purposes
    def to_s(sheet = nil)
      sheet_for(sheet).cells.inspect
    end

    # returns the row,col values of the labelled cell
    # (nil,nil) if label is not defined
    def label(name)
      labels = workbook.defined_names
      return [nil, nil, nil] if labels.empty? || !labels.key?(name)

      [labels[name].row, labels[name].col, labels[name].sheet]
    end

    # Returns an array which all labels. Each element is an array with
    # [labelname, [row,col,sheetname]]
    def labels
      @labels ||= workbook.defined_names.map do |name, label|
        [
          name,
          [label.row, label.col, label.sheet]
        ]
      end
    end

    def hyperlink?(row, col, sheet = nil)
      !!hyperlink(row, col, sheet)
    end

    # returns the hyperlink at (row/col)
    # nil if there is no hyperlink
    def hyperlink(row, col, sheet = nil)
      key = normalize(row, col)
      sheet_for(sheet).hyperlinks[key]
    end

    # returns the comment at (row/col)
    # nil if there is no comment
    def comment(row, col, sheet = nil)
      key = normalize(row, col)
      sheet_for(sheet).comments[key]
    end

    # true, if there is a comment
    def comment?(row, col, sheet = nil)
      !!comment(row, col, sheet)
    end

    def comments(sheet = nil)
      sheet_for(sheet).comments.map do |(x, y), comment|
        [x, y, comment]
      end
    end

    # Yield an array of Excelx::Cell
    # Takes options for sheet, pad_cells, and max_rows
    def each_row_streaming(options = {})
      sheet = sheet_for(options.delete(:sheet))
      if block_given?
        sheet.each_row(options) { |row| yield row }
      else
        sheet.to_enum(:each_row, options)
      end
    end

    private

    def clean_sheet(sheet)
      @sheets_by_name[sheet].cells.each_pair do |coord, value|
        next unless value.value.is_a?(::String)

        @sheets_by_name[sheet].cells[coord].value = sanitize_value(value.value)
      end

      @cleaned[sheet] = true
    end

    # Internal: extracts the worksheet_ids from the workbook.xml file. xlsx
    #           documents require a workbook.xml file, so a if the file is missing
    #           it is not a valid xlsx file. In these cases, an ArgumentError is
    #           raised.
    #
    # wb - a Zip::Entry for the workbook.xml file.
    # path - A String for Zip::Entry's destination path.
    #
    # Examples
    #
    #   extract_worksheet_ids(<Zip::Entry>, 'tmpdir/roo_workbook.xml')
    #   # => ["rId1", "rId2", "rId3"]
    #
    # Returns an Array of Strings.
    def extract_worksheet_ids(entries, path)
      wb = entries.find { |e| e.name[/workbook.xml$/] }
      fail ArgumentError 'missing required workbook file' if wb.nil?

      wb.extract(path)
      workbook_doc = Roo::Utils.load_xml(path).remove_namespaces!
      workbook_doc.xpath('//sheet').map { |s| s['id'] }
    end

    # Internal
    #
    # wb_rels - A Zip::Entry for the workbook.xml.rels file.
    # path - A String for the Zip::Entry's destination path.
    #
    # Examples
    #
    #   extract_worksheets(<Zip::Entry>, 'tmpdir/roo_workbook.xml.rels')
    #   # => {
    #         "rId1"=>"worksheets/sheet1.xml",
    #         "rId2"=>"worksheets/sheet2.xml",
    #         "rId3"=>"worksheets/sheet3.xml"
    #        }
    #
    # Returns a Hash.
    def extract_worksheet_rels(entries, path)
      wb_rels = entries.find { |e| e.name[/workbook.xml.rels$/] }
      fail ArgumentError 'missing required workbook file' if wb_rels.nil?

      wb_rels.extract(path)
      rels_doc = Roo::Utils.load_xml(path).remove_namespaces!

      relationships = rels_doc.xpath('//Relationship').select do |relationship|
        worksheet_types.include? relationship['Type']
      end

      relationships.each_with_object({}) do |relationship, hash|
        hash[relationship['Id']] = relationship['Target']
      end
    end

    # Extracts the sheets in order, but it will ignore sheets that are not
    # worksheets.
    def extract_sheets_in_order(entries, sheet_ids, sheets, tmpdir)
      (sheet_ids & sheets.keys).each_with_index do |id, i|
        name = sheets[id]
        entry = entries.find { |e| "/#{e.name}" =~ /#{name}$/ }
        path = "#{tmpdir}/roo_sheet#{i + 1}"
        sheet_files << path
        @sheet_files << path
        entry.extract(path)
      end
    end

    def extract_images(entries, tmpdir)
      img_entries = entries.select { |e| e.name[/media\/image([0-9]+)/] }
      img_entries.each do |entry|
        path = "#{@tmpdir}/roo#{entry.name.gsub(/xl\/|\//, "_")}"
        image_files << path
        entry.extract(path)
      end
    end

    # Extracts all needed files from the zip file
    def process_zipfile(zipfilename_or_stream)
      @sheet_files = []

      unless is_stream?(zipfilename_or_stream)
        zip_file = Zip::File.open(zipfilename_or_stream)
      else
        zip_file = Zip::CentralDirectory.new
        zip_file.read_from_stream zipfilename_or_stream
      end

      process_zipfile_entries zip_file.to_a.sort_by(&:name)
    end

    def process_zipfile_entries(entries)
      # NOTE: When Google or Numbers 3.1 exports to xlsx, the worksheet filenames
      #       are not in order. With Numbers 3.1, the first sheet is always
      #       sheet.xml, not sheet1.xml. With Google, the order of the worksheets is
      #       independent of a worksheet's filename (i.e. sheet6.xml can be the
      #       first worksheet).
      #
      #       workbook.xml lists the correct order of worksheets and
      #       workbook.xml.rels lists the filenames for those worksheets.
      #
      #       workbook.xml:
      #         <sheet state="visible" name="IS" sheetId="1" r:id="rId3"/>
      #         <sheet state="visible" name="BS" sheetId="2" r:id="rId4"/>
      #       workbook.xml.rel:
      #         <Relationship Id="rId4" Target="worksheets/sheet5.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"/>
      #         <Relationship Id="rId3" Target="worksheets/sheet4.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"/>
      sheet_ids = extract_worksheet_ids(entries, "#{@tmpdir}/roo_workbook.xml")
      sheets = extract_worksheet_rels(entries, "#{@tmpdir}/roo_workbook.xml.rels")
      extract_sheets_in_order(entries, sheet_ids, sheets, @tmpdir)
      extract_images(entries, @tmpdir)

      entries.each do |entry|
        path =
        case entry.name.downcase
        when /sharedstrings.xml$/
          "#{@tmpdir}/roo_sharedStrings.xml"
        when /styles.xml$/
          "#{@tmpdir}/roo_styles.xml"
        when /comments([0-9]+).xml$/
          # FIXME: Most of the time, The order of the comment files are the same
          #       the sheet order, i.e. sheet1.xml's comments are in comments1.xml.
          #       In some situations, this isn't true. The true location of a
          #       sheet's comment file is in the sheet1.xml.rels file. SEE
          #       ECMA-376 12.3.3 in "Ecma Office Open XML Part 1".
          nr = Regexp.last_match[1].to_i
          comments_files[nr - 1] = "#{@tmpdir}/roo_comments#{nr}"
        when %r{chartsheets/_rels/sheet([0-9]+).xml.rels$}
          # NOTE: Chart sheet relationship files were interfering with
          #       worksheets.
          nil
        when /sheet([0-9]+).xml.rels$/
          # FIXME: Roo seems to use sheet[\d].xml.rels for hyperlinks only, but
          #        it also stores the location for sharedStrings, comments,
          #        drawings, etc.
          nr = Regexp.last_match[1].to_i
          rels_files[nr - 1] = "#{@tmpdir}/roo_rels#{nr}"
        when /drawing([0-9]+).xml.rels$/
          # Extracting drawing relationships to make images lists for each sheet
          nr = Regexp.last_match[1].to_i
          image_rels[nr - 1] = "#{@tmpdir}/roo_image_rels#{nr}"
        end

        entry.extract(path) if path
      end
    end

    def safe_send(object, method, *args)
      object.send(method, *args) if object&.respond_to?(method)
    end

    def worksheet_types
      [
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet', # OOXML Transitional
        'http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet' # OOXML Strict
      ]
    end
  end
end
