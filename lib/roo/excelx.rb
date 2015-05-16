require 'nokogiri'
require 'zip/filesystem'
require 'roo/link'
require 'roo/utils'

module Roo
  class Excelx < Roo::Base
    require 'roo/excelx/workbook'
    require 'roo/excelx/shared_strings'
    require 'roo/excelx/styles'
    require 'roo/excelx/cell'
    require 'roo/excelx/sheet'
    require 'roo/excelx/relationships'
    require 'roo/excelx/comments'
    require 'roo/excelx/sheet_doc'
    
    module Format
      EXCEPTIONAL_FORMATS = {
        'h:mm am/pm' => :date,
        'h:mm:ss am/pm' => :date
      }

      STANDARD_FORMATS = {
        0 => 'General'.freeze,
        1 => '0'.freeze,
        2 => '0.00'.freeze,
        3 => '#,##0'.freeze,
        4 => '#,##0.00'.freeze,
        9 => '0%'.freeze,
        10 => '0.00%'.freeze,
        11 => '0.00E+00'.freeze,
        12 => '# ?/?'.freeze,
        13 => '# ??/??'.freeze,
        14 => 'mm-dd-yy'.freeze,
        15 => 'd-mmm-yy'.freeze,
        16 => 'd-mmm'.freeze,
        17 => 'mmm-yy'.freeze,
        18 => 'h:mm AM/PM'.freeze,
        19 => 'h:mm:ss AM/PM'.freeze,
        20 => 'h:mm'.freeze,
        21 => 'h:mm:ss'.freeze,
        22 => 'm/d/yy h:mm'.freeze,
        37 => '#,##0 ;(#,##0)'.freeze,
        38 => '#,##0 ;[Red](#,##0)'.freeze,
        39 => '#,##0.00;(#,##0.00)'.freeze,
        40 => '#,##0.00;[Red](#,##0.00)'.freeze,
        45 => 'mm:ss'.freeze,
        46 => '[h]:mm:ss'.freeze,
        47 => 'mmss.0'.freeze,
        48 => '##0.0E+0'.freeze,
        49 => '@'.freeze
      }

      def to_type(format)
        format = format.to_s.downcase
        if (type = EXCEPTIONAL_FORMATS[format])
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

      unless is_stream?(filename_or_stream)
        file_type_check(filename_or_stream, '.xlsx', 'an Excel-xlsx', file_warning, packed)
        basename = File.basename(filename_or_stream)
      end

      @tmpdir = make_tmpdir(basename, options[:tmpdir_root])
      @filename = local_filename(filename_or_stream, @tmpdir, packed)
      @comments_files = []
      @rels_files = []
      process_zipfile(@filename || filename_or_stream)

      @sheet_names = workbook.sheets.map do |sheet|
        unless options[:only_visible_sheets] && sheet['state'] == 'hidden'
          sheet['name']
        end
      end.compact
      @sheets = []
      @sheets_by_name = Hash[@sheet_names.map.with_index do |sheet_name, n|
        @sheets[n] = Sheet.new(sheet_name, @rels_files[n], @sheet_files[n], @comments_files[n], styles, shared_strings, workbook, sheet_options)
        [sheet_name, @sheets[n]]
      end]

      if cell_max
        cell_count = ::Roo::Utils.num_cells_in_range(sheet_for(options.delete(:sheet)).dimensions)
        raise ExceedsMaxError.new("Excel file exceeds cell maximum: #{cell_count} > #{cell_max}") if cell_count > cell_max
      end

      super
    rescue => e # clean up any temp files, but only if an error was raised
      close
      raise e
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
      @sheets_by_name[sheet]
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
      sheet_for(sheet).cells[key] = Cell.new(value, cell_type, nil, cell_type, value, nil, nil, nil, Cell::Coordinate.new(row, col))
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
      safe_send(sheet_for(sheet).cells[key], :excelx_type)
    end

    # returns the internal value of an excelx cell
    # Note: this is only available within the Excelx class
    def excelx_value(row, col, sheet = nil)
      key = normalize(row, col)
      safe_send(sheet_for(sheet).cells[key], :excelx_value)
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
      !cell || !cell.value || (cell.type == :string && cell.value.empty?) \
      || (row < sheet.first_row || row > sheet.last_row || col < sheet.first_column || col > sheet.last_column)
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
      sheet_for(options.delete(:sheet)).each_row(options) { |row| yield row }
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
      workbook_doc.xpath('//sheet').map { |s| s.attributes['id'].value }
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
      worksheet_type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'

      relationships = rels_doc.xpath('//Relationship').select do |relationship|
        relationship.attributes['Type'].value == worksheet_type
      end

      relationships.inject({}) do |hash, relationship|
        attributes = relationship.attributes
        id = attributes['Id']
        hash[id.value] = attributes['Target'].value
        hash
      end
    end

    def extract_sheets_in_order(entries, sheet_ids, sheets, tmpdir)
      sheet_ids.each_with_index do |id, i|
        name = sheets[id]
        entry = entries.find { |e| e.name =~ /#{name}$/ }
        path = "#{tmpdir}/roo_sheet#{i + 1}"
        @sheet_files << path
        entry.extract(path)
      end
    end

    # Extracts all needed files from the zip file
    def process_zipfile(zipfilename_or_stream)
      @sheet_files = []

      unless is_stream?(zipfilename_or_stream)
        process_zipfile_entries Zip::File.open(zipfilename_or_stream).to_a.sort_by(&:name)
      else
        stream = Zip::InputStream.open zipfilename_or_stream
        begin
          entries = []
          while (entry = stream.get_next_entry)
            entries << entry
          end
          process_zipfile_entries entries
        ensure
          stream.close
        end
      end
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
          @comments_files[nr - 1] = "#{@tmpdir}/roo_comments#{nr}"
        when /sheet([0-9]+).xml.rels$/
          # FIXME: Roo seems to use sheet[\d].xml.rels for hyperlinks only, but
          #        it also stores the location for sharedStrings, comments,
          #        drawings, etc.
          nr = Regexp.last_match[1].to_i
          @rels_files[nr - 1] = "#{@tmpdir}/roo_rels#{nr}"
        end

        entry.extract(path) if path
      end
    end

    def styles
      @styles ||= Styles.new(File.join(@tmpdir, 'roo_styles.xml'))
    end

    def shared_strings
      @shared_strings ||= SharedStrings.new(File.join(@tmpdir, 'roo_sharedStrings.xml'))
    end

    def workbook
      @workbook ||= Workbook.new(File.join(@tmpdir, 'roo_workbook.xml'))
    end

    def safe_send(object, method, *args)
      object.send(method, *args) if object && object.respond_to?(method)
    end
  end
end
