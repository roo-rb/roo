# frozen_string_literal: true

require "csv"
require "time"

# The CSV class can read csv files (must be separated with commas) which then
# can be handled like spreadsheets. This means you can access cells like A5
# within these files.
# The CSV class provides only string objects. If you want conversions to other
# types you have to do it yourself.
#
# You can pass options to the underlying CSV parse operation, via the
# :csv_options option.
module Roo
  class CSV < Roo::Base
    attr_reader :filename

    # Returns an array with the names of the sheets. In CSV class there is only
    # one dummy sheet, because a csv file cannot have more than one sheet.
    def sheets
      ["default"]
    end

    def cell(row, col, sheet = nil)
      sheet ||= default_sheet
      read_cells(sheet)
      @cell[normalize(row, col)]
    end

    def celltype(row, col, sheet = nil)
      sheet ||= default_sheet
      read_cells(sheet)
      @cell_type[normalize(row, col)]
    end

    def cell_postprocessing(_row, _col, value)
      value
    end

    def csv_options
      @options[:csv_options] || {}
    end

    def set_value(row, col, value, _sheet)
      @cell[[row, col]] = value
    end

    def set_type(row, col, type, _sheet)
      @cell_type[[row, col]] = type
    end

    private

    TYPE_MAP = {
      String => :string,
      Float => :float,
      Date => :date,
      DateTime => :datetime,
    }

    def celltype_class(value)
      TYPE_MAP[value.class]
    end

    def read_cells(sheet = default_sheet)
      sheet ||= default_sheet
      return if @cells_read[sheet]
      row_num = 0
      max_col_num = 0

      each_row csv_options do |row|
        row_num += 1
        col_num = 0

        row.each do |elem|
          col_num += 1
          coordinate = [row_num, col_num]
          @cell[coordinate] = elem
          @cell_type[coordinate] = celltype_class(elem)
        end

        max_col_num = col_num if col_num > max_col_num
      end

      set_row_count(sheet, row_num)
      set_column_count(sheet, max_col_num)
      @cells_read[sheet] = true
    end

    def each_row(options, &block)
      if uri?(filename)
        each_row_using_tempdir(options, &block)
      elsif is_stream?(filename_or_stream)
        ::CSV.new(filename_or_stream, options).each(&block)
      else
        ::CSV.foreach(filename, options, &block)
      end
    end

    def each_row_using_tempdir(options, &block)
      ::Dir.mktmpdir(Roo::TEMP_PREFIX, ENV["ROO_TMP"]) do |tmpdir|
        tmp_filename = download_uri(filename, tmpdir)
        ::CSV.foreach(tmp_filename, options, &block)
      end
    end

    def set_row_count(sheet, last_row)
      @first_row[sheet] = 1
      @last_row[sheet] = last_row
      @last_row[sheet] = @first_row[sheet] if @last_row[sheet].zero?

      nil
    end

    def set_column_count(sheet, last_col)
      @first_column[sheet] = 1
      @last_column[sheet] = last_col
      @last_column[sheet] = @first_column[sheet] if @last_column[sheet].zero?

      nil
    end

    def clean_sheet(sheet)
      read_cells(sheet)

      @cell.each_pair do |coord, value|
        @cell[coord] = sanitize_value(value) if value.is_a?(::String)
      end

      @cleaned[sheet] = true
    end

    alias_method :filename_or_stream, :filename
  end
end
