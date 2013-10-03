require 'csv'
require 'time'

# The CSV class can read csv files (must be separated with commas) which then
# can be handled like spreadsheets. This means you can access cells like A5
# within these files.
# The CSV class provides only string objects. If you want conversions to other
# types you have to do it yourself.
#
# You can pass options to the underlying CSV parse operation, via the
# :csv_options option.
#

class Roo::CSV < Roo::Base
  def initialize(filename, options = {})
    super
  end

  attr_reader :filename

  # Returns an array with the names of the sheets. In CSV class there is only
  # one dummy sheet, because a csv file cannot have more than one sheet.
  def sheets
    ['default']
  end

  def cell(row, col, sheet=nil)
    sheet ||= @default_sheet
    read_cells(sheet)
    @cell[normalize(row,col)]
  end

  def celltype(row, col, sheet=nil)
    sheet ||= @default_sheet
    read_cells(sheet)
    @cell_type[normalize(row,col)]
  end

  def cell_postprocessing(row,col,value)
    value
  end

  def csv_options
    @options[:csv_options] || {}
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

  def each_row(options, &block)
    if uri?(filename)
      make_tmpdir do |tmpdir|
        tmp_filename = download_uri(filename, tmpdir)
        CSV.foreach(tmp_filename, options, &block)
      end
    else
      CSV.foreach(filename, options, &block)
    end
  end

  def read_cells(sheet=nil)
    sheet ||= @default_sheet
    return if @cells_read[sheet]
    @first_row[sheet] = 1
    @last_row[sheet] = 0
    @first_column[sheet] = 1
    @last_column[sheet] = 1
    rownum = 1
    each_row csv_options do |row|
      row.each_with_index do |elem,i|
        @cell[[rownum,i+1]] = cell_postprocessing rownum,i+1, elem
        @cell_type[[rownum,i+1]] = celltype_class @cell[[rownum,i+1]]
        if i+1 > @last_column[sheet]
          @last_column[sheet] += 1
        end
      end
      rownum += 1
      @last_row[sheet] += 1
    end
    @cells_read[sheet] = true
    #-- adjust @first_row if neccessary
    while !row(@first_row[sheet]).any? and @first_row[sheet] < @last_row[sheet]
      @first_row[sheet] += 1
    end
    #-- adjust @last_row if neccessary
    while !row(@last_row[sheet]).any? and @last_row[sheet] and
        @last_row[sheet] > @first_row[sheet]
      @last_row[sheet] -= 1
    end
    #-- adjust @first_column if neccessary
    while !column(@first_column[sheet]).any? and
          @first_column[sheet] and
          @first_column[sheet] < @last_column[sheet]
      @first_column[sheet] += 1
    end
    #-- adjust @last_column if neccessary
    while !column(@last_column[sheet]).any? and
          @last_column[sheet] and
          @last_column[sheet] > @first_column[sheet]
      @last_column[sheet] -= 1
    end
  end
end
