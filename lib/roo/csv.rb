require 'rubygems'
require 'csv'
require 'time'
require 'iconv'

IC = Iconv.new('UTF-8//IGNORE', 'UTF-8')   # set up the ic accessor to hold the Iconv force encoding object

# The Csv class can read csv files (must be separated with commas) which then
# can be handled like spreadsheets. This means you can access cells like A5
# within these files.
# The Csv class provides only string objects. If you want conversions to other
# types you have to do it yourself.

class Roo::Csv < Roo::GenericSpreadsheet
  def initialize(filename, packed=nil, file_warning=:error, tmpdir=nil, options=Hash.new)
    @filename = filename
    @cell = Hash.new
    @cell_type = Hash.new
    @cells_read = Hash.new
    @first_row = Hash.new
    @last_row = Hash.new
    @first_column = Hash.new
    @last_column = Hash.new
    @options = options
  end

  # Returns an array with the names of the sheets. In Csv class there is only
  # one dummy sheet, because a csv file cannot have more than one sheet.
  def sheets
    ['default']
  end

  def cell(row, col, sheet=nil)
    sheet ||= @default_sheet
    read_cells(sheet) unless @cells_read[sheet]
    row,col = normalize(row,col)
    @cell[[row,col]]
  end

  def celltype(row, col, sheet=nil)
    sheet ||= @default_sheet
    read_cells(sheet) unless @cells_read[sheet]
    row,col = normalize(row,col)
    @cell_type[[row,col]]
  end

  def cell_postprocessing(row,col,value)
    value
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

  # Use iconv force encoding before parsing the line using csv
  def parse_line(line)
    line = (IC.iconv(line + ' ')[0..-2]).gsub /"/, ''
    CSV.parse_line(line, @options)
  end

  def read_cells(sheet=nil)
    sheet ||= @default_sheet
    @cell_type = {} unless @cell_type
    @cell = {} unless @cell
    @first_row[sheet] = 1
    @last_row[sheet] = 0
    @first_column[sheet] = 1
    @last_column[sheet] = 1
    rownum = 1

    f = File.open(@filename)
    f.each_line do |line|
      row = parse_line line
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
    loop do
      if !row(@first_row[sheet]).any? and @first_row[sheet] < @last_row[sheet]
        @first_row[sheet] += 1
      else
        break
      end
    end
    #-- adjust @last_row if neccessary
    loop do
      if !row(@last_row[sheet]).any? and @last_row[sheet] and
          @last_row[sheet] > @first_row[sheet]
        @last_row[sheet] -= 1
      else
        break
      end
    end
    #-- adjust @first_column if neccessary
    loop do
      if !column(@first_column[sheet]).any? and
          @first_column[sheet] and
          @first_column[sheet] < @last_column[sheet]
        @first_column[sheet] += 1
      else
        break
      end
    end
    #-- adjust @last_column if neccessary
    loop do
      if !column(@last_column[sheet]).any? and
          @last_column[sheet] and
          @last_column[sheet] > @first_column[sheet]
        @last_column[sheet] -= 1
      else
        break
      end
    end
  end
end # class Csv
