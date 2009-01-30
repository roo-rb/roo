require 'rubygems'
require 'builder'

# Base class for all other types of spreadsheets
class GenericSpreadsheet

  attr_reader :default_sheet

  # sets the line with attribute names (default: 1)
  attr_accessor :header_line

  def initialize
  end

  # set the working sheet in the document
  def default_sheet=(sheet)
    if sheet.kind_of? Fixnum
      if sheet >= 0 and sheet <= sheets.length
        sheet = self.sheets[sheet-1]
      else
        raise RangeError
      end
    elsif sheet.kind_of?(String)
      raise RangeError if ! self.sheets.include?(sheet)
    else
      raise TypeError, "what are you trying to set as default sheet?"
    end
    @default_sheet = sheet
    check_default_sheet
    @first_row[sheet] = @last_row[sheet] = @first_column[sheet] = @last_column[sheet] = nil
    @cells_read[sheet] = false
  end

  # first non-empty column as a letter
  def first_column_as_letter(sheet=nil)
    GenericSpreadsheet.number_to_letter(first_column(sheet))
  end

  # last non-empty column as a letter
  def last_column_as_letter(sheet=nil)
    GenericSpreadsheet.number_to_letter(last_column(sheet))
  end

  # returns the number of the first non-empty row
  def first_row(sheet=nil)
    if sheet == nil
      sheet = @default_sheet
    end
    read_cells(sheet) unless @cells_read[sheet]
    if @first_row[sheet]
      return @first_row[sheet]
    end
    impossible_value = 999_999 # more than a spreadsheet can hold
    result = impossible_value
    @cell[sheet].each_pair {|key,value|
      y,x = key # _to_string(key).split(',')
      y = y.to_i
      result = [result, y].min if value
    } if @cell[sheet]
    result = nil if result == impossible_value
    @first_row[sheet] = result
    result
  end

  # returns the number of the last non-empty row
  def last_row(sheet=nil)
    sheet = @default_sheet unless sheet
    read_cells(sheet) unless @cells_read[sheet]
    if @last_row[sheet]
      return @last_row[sheet]
    end
    impossible_value = 0
    result = impossible_value
    @cell[sheet].each_pair {|key,value|
      y,x = key # _to_string(key).split(',')
      y = y.to_i
      result = [result, y].max if value
    } if @cell[sheet]
    result = nil if result == impossible_value
    @last_row[sheet] = result
    result
  end

  # returns the number of the first non-empty column
  def first_column(sheet=nil)
    if sheet == nil
      sheet = @default_sheet
    end
    read_cells(sheet) unless @cells_read[sheet]
    if @first_column[sheet]
      return @first_column[sheet]
    end
    impossible_value = 999_999 # more than a spreadsheet can hold
    result = impossible_value
    @cell[sheet].each_pair {|key,value|
      y,x = key # _to_string(key).split(',')
      x = x # .to_i
      result = [result, x].min if value
    } if @cell[sheet]
    result = nil if result == impossible_value
    @first_column[sheet] = result
    result
  end

  # returns the number of the last non-empty column
  def last_column(sheet=nil)
    sheet = @default_sheet unless sheet
    read_cells(sheet) unless @cells_read[sheet]
    if @last_column[sheet]
      return @last_column[sheet]
    end
    impossible_value = 0
    result = impossible_value
    @cell[sheet].each_pair {|key,value|
      y,x = key # _to_string(key).split(',')
      x = x.to_i
      result = [result, x].max if value
    } if @cell[sheet]
    result = nil if result == impossible_value
    @last_column[sheet] = result
    result
  end

  # returns a rectangular area (default: all cells) as yaml-output
  # you can add additional attributes with the prefix parameter like:
  # oo.to_yaml({"file"=>"flightdata_2007-06-26", "sheet" => "1"})
  def to_yaml(prefix={}, from_row=nil, from_column=nil, to_row=nil, to_column=nil,sheet=nil)
    sheet = @default_sheet unless sheet
    result = "--- \n"
    (from_row||first_row(sheet)).upto(to_row||last_row(sheet)) do |row|
      (from_column||first_column(sheet)).upto(to_column||last_column(sheet)) do |col|
        unless empty?(row,col,sheet)
          result << "cell_#{row}_#{col}: \n"
          prefix.each {|k,v|
            result << "  #{k}: #{v} \n"
          }
          result << "  row: #{row} \n"
          result << "  col: #{col} \n"
          result << "  celltype: #{self.celltype(row,col,sheet)} \n"
          if self.celltype(row,col,sheet) == :time
            result << "  value: #{GenericSpreadsheet.integer_to_timestring( self.cell(row,col,sheet))} \n"
          else
            result << "  value: #{self.cell(row,col,sheet)} \n"
          end
        end
      end
    end
    result
  end

  # write the current spreadsheet to stdout or into a file
  def to_csv(filename=nil,sheet=nil)
    sheet = @default_sheet unless sheet
    if filename
      file = File.open(filename,"w") # do |file|
      write_csv_content(file,sheet)
      file.close
    else
      write_csv_content(STDOUT,sheet)
    end
    true
  end

  # find a row either by row number or a condition
  # Caution: this works only within the default sheet -> set default_sheet before you call this method
  # (experimental. see examples in the test_roo.rb file)
  def find(*args) # :nodoc
    result_array = false
    args.each {|arg,val|
      if arg.class == Hash
        arg.each { |hkey,hval|
          if hkey == :array and hval == true
            result_array = true
          end
        }
      end
    }
    column_with = {}
    1.upto(last_column) do |col|
      column_with[cell(@header_line,col)] = col
    end
    result = Array.new
    #-- id
    if args[0].class == Fixnum
      rownum = args[0]
      if @header_line
        tmp = {}
      else
        tmp = []
      end
      1.upto(self.row(rownum).size) {|j|
        x = ''
        column_with.each { |key,val|
          if val == j
            x = key
          end
        }
        if @header_line
          tmp[x] = cell(rownum,j)
        else
          tmp[j-1] = cell(rownum,j)
        end

      }
      if @header_line
        result = [ tmp ]
      else
        result = tmp
      end
      #-- :all
    elsif args[0] == :all
      if args[1].class == Hash
        args[1].each {|key,val|
          if key == :conditions
            column_with = {}
            1.upto(last_column) do |col|
              column_with[cell(@header_line,col)] = col
            end
            conditions = val
            first_row.upto(last_row) do |i|
              # are all conditions met?
              found = 1
              conditions.each { |key,val|
                if cell(i,column_with[key]) == val
                  found *= 1
                else
                  found *= 0
                end
              }
              if found > 0
                tmp = {}
                1.upto(self.row(i).size) {|j|
                  x = ''
                  column_with.each { |key,val|
                    if val == j
                      x = key
                    end
                  }
                  tmp[x] = cell(i,j)
                }
                if result_array
                  result << self.row(i)
                else
                  result << tmp
                end
              end
            end
          end # :conditions
        }
      end
    end
    result
  end

  # returns all values in this row as an array
  # row numbers are 1,2,3,... like in the spreadsheet
  def row(rownumber,sheet=nil)
    sheet = @default_sheet unless sheet
    read_cells(sheet) unless @cells_read[sheet]
    result = []
    tmp_arr = []
    @cell[sheet].each_pair {|key,value|
      y,x = key # _to_string(key).split(',')
      x = x.to_i
      y = y.to_i
      if y == rownumber
        tmp_arr[x] = value
      end
    }
    result = tmp_arr[1..-1]
    while result && result[-1] == nil
      result = result[0..-2]
    end
    result
  end

  # returns all values in this column as an array
  # column numbers are 1,2,3,... like in the spreadsheet
  def column(columnnumber,sheet=nil)
    if columnnumber.class == String
      columnnumber = Excel.letter_to_number(columnnumber)
    end
    sheet = @default_sheet unless sheet
    read_cells(sheet) unless @cells_read[sheet]
    result = []
    first_row(sheet).upto(last_row(sheet)) do |row|
      result << cell(row,columnnumber,sheet)
    end
    result
  end

  # reopens and read a spreadsheet document
  def reload
    ds = @default_sheet
    initialize(@filename) if self.class == Openoffice or
      self.class == Excel
    initialize(@spreadsheetkey,@user,@password) if self.class == Google
    self.default_sheet = ds
    #@first_row = @last_row = @first_column = @last_column = nil
  end

  # true if cell is empty
  def empty?(row, col, sheet=nil)
    sheet = @default_sheet unless sheet
    read_cells(sheet) unless @cells_read[sheet] or self.class == Excel
    row,col = normalize(row,col)
    return true unless cell(row, col, sheet)
    return true if celltype(row, col, sheet) == :string && cell(row, col, sheet).empty?
    return true if row < first_row(sheet) || row > last_row(sheet) || col < first_column(sheet) || col > last_column(sheet)
    false
  end

  # recursively removes the current temporary directory
  # this is only needed if you work with zipped files or files via the web
  def remove_tmp
    if File.exists?(@tmpdir)
      FileUtils::rm_r(@tmpdir)
    end
  end

  # Returns information of the spreadsheet document and all sheets within
  # this document.
  def info
    result = "File: #{File.basename(@filename)}\n"+
      "Number of sheets: #{sheets.size}\n"+
      "Sheets: #{sheets.map{|sheet| sheet+", "}.to_s[0..-3]}\n"
    n = 1
    sheets.each {|sheet|
      self.default_sheet = sheet
      result << "Sheet " + n.to_s + ":\n"
      unless first_row
        result << "  - empty -"
      else
        result << "  First row: #{first_row}\n"
        result << "  Last row: #{last_row}\n"
        result << "  First column: #{GenericSpreadsheet.number_to_letter(first_column)}\n"
        result << "  Last column: #{GenericSpreadsheet.number_to_letter(last_column)}"
      end
      result << "\n" if sheet != sheets.last
      n += 1
    }
    result
  end

  def to_xml
    xml_document = ''
    xml = Builder::XmlMarkup.new(:target => xml_document, :indent => 2)
    xml.instruct! :xml, :version =>"1.0", :encoding => "utf-8"
    xml.spreadsheet {
      self.sheets.each do |sheet|
        self.default_sheet = sheet
        xml.sheet(:name => sheet) { |x|
          if first_row and last_row and first_column and last_column
            # sonst gibt es Fehler bei leeren Blaettern
            first_row.upto(last_row) do |row|
              first_column.upto(last_column) do |col|
                unless empty?(row,col)
                  x.cell(cell(row,col),
                    :row =>row,
                    :column => col,
                    :type => celltype(row,col))
                end
              end
            end
          end
        }
      end
    }
    xml_document
  end

  protected

  def file_type_check(filename, ext, name)
    new_expression = {
      '.ods' => 'Openoffice.new',
      '.xls' => 'Excel.new',
      '.xlsx' => 'Excelx.new',
    }
    case ext
    when '.ods', '.xls', '.xlsx'
      correct_class = "use #{new_expression[ext]} to handle #{ext} spreadsheet files"
    else
      raise "unknown file type: #{ext}"
    end
    if File.extname(filename).downcase != ext
      case @file_warning
      when :error
        warn correct_class
        raise TypeError, "#{filename} is not #{name} file"
      when :warning
        warn "are you sure, this is #{name} spreadsheet file?"
        warn correct_class
      when :ignore
        # ignore
      else
        raise "#{@file_warning} illegal state of file_warning"
      end
    end
  end

  # konvertiert einen Key in der Form "12,45" (=row,column) in
  # ein Array mit numerischen Werten ([12,45])
  # Diese Methode ist eine temp. Loesung, um zu erforschen, ob der 
  # Zugriff mit numerischen Keys schneller ist.
  def key_to_num(str)
    r,c = str.split(',')
    r = r.to_i
    c = c.to_i
    [r,c]
  end

  # siehe: key_to_num
  def key_to_string(arr)
    "#{arr[0]},#{arr[1]}"
  end 

  private

  # converts cell coordinate to numeric values of row,col
  def normalize(row,col)
    if row.class == String
      if col.class == Fixnum
        # ('A',1):
        # ('B', 5) -> (5, 2)
        row, col = col, row
      else
        raise ArgumentError
      end
    end
    if col.class == String
      col = GenericSpreadsheet.letter_to_number(col)
    end
    return row,col
  end

  #  def open_from_uri(uri)
  #    require 'open-uri' ;
  #    tempfilename = File.join(@tmpdir, File.basename(uri))
  #    f = File.open(tempfilename,"wb")
  #    begin
  #      open(uri) do |net|
  #        f.write(net.read)
  #      end
  #    rescue
  #      raise "could not open #{uri}"
  #    end
  #    f.close
  #    File.join(@tmpdir, File.basename(uri))
  #  end

  #  OpenURI::HTTPError
  #  def open_from_uri(uri)
  #    require 'open-uri'
  #    #existiert URL?
  #    r = Net::HTTP.get_response(URI.parse(uri))
  #    raise "URL nicht verfuegbar" unless r.is_a? Net::HTTPOK
  #    tempfilename = File.join(@tmpdir, File.basename(uri))
  #    f = File.open(tempfilename,"wb")
  #    open(uri) do |net|
  #      f.write(net.read)
  #    end
  #    #   rescue
  #    #    raise "could not open #{uri}"
  #    # end
  #    f.close
  #    File.join(@tmpdir, File.basename(uri))
  #  end
  
  def open_from_uri(uri)
    require 'open-uri'
    response = ''
    begin
      open(uri, "User-Agent" => "Ruby/#{RUBY_VERSION}") { |net| 
        response = net.read 
        tempfilename = File.join(@tmpdir, File.basename(uri))
        f = File.open(tempfilename,"wb")
        f.write(response)
        f.close
      }
    rescue OpenURI::HTTPError
      raise "could not open #{uri}"
    end
    File.join(@tmpdir, File.basename(uri))
  end
  
  def open_from_stream(stream)
    tempfilename = File.join(@tmpdir, "spreadsheet")
    f = File.open(tempfilename,"wb")
    f.write(stream[7..-1])
    f.close
    File.join(@tmpdir, "spreadsheet")
  end

  # convert a number to something like 'AB' (1 => 'A', 2 => 'B', ...)
  def self.number_to_letter(n)
    letters=""
    while n > 0
      num = n%26
      letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"[num-1,1] + letters
      n = n.div(26)
    end
    letters
  end

  # convert letters like 'AB' to a number ('A' => 1, 'B' => 2, ...)
  def self.letter_to_number(letters)
    result = 0
    while letters && letters.length > 0
      character = letters[0,1].upcase
      num = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".index(character)
      raise ArgumentError, "invalid column character '#{letters[0,1]}'" if num == nil
      num += 1
      result = result * 26 + num
      letters = letters[1..-1]
    end
    result
  end

  def unzip(filename)
    ret = nil
    Zip::ZipFile.open(filename) do |zip|
      ret = process_zipfile_packed zip
    end
    ret
  end

  # check if default_sheet was set and exists in sheets-array
  def check_default_sheet
    sheet_found = false
    raise ArgumentError, "Error: default_sheet not set" if @default_sheet == nil
    if sheets.index(@default_sheet)
      sheet_found = true
    end
    if ! sheet_found
      raise RangeError, "sheet '#{@default_sheet}' not found"
    end
    #raise ArgumentError, "Error: default_sheet not set" if @default_sheet == nil
  end

  def process_zipfile_packed(zip, path='')
    ret=nil
    if zip.file.file? path
      # extract and return filename
      @tmpdir = "oo_"+$$.to_s
      unless File.exists?(@tmpdir)
        FileUtils::mkdir(@tmpdir)
      end
      file = File.open(File.join(@tmpdir, path),"wb")
      file.write(zip.read(path))
      file.close
      return File.join(@tmpdir, path)
    else
      unless path.empty?
        path += '/'
      end
      zip.dir.foreach(path) do |filename|
        ret = process_zipfile_packed(zip, path + filename)
      end
    end
    ret
  end

  def write_csv_content(file=nil,sheet=nil)
    file = STDOUT unless file
    if first_row(sheet) # sheet is not empty
      # first_row(sheet).upto(last_row(sheet)) do |row|
      1.upto(last_row(sheet)) do |row|
        1.upto(last_column(sheet)) do |col|
          file.print(",") if col > 1
          onecell = cell(row,col,sheet)
          onecelltype = celltype(row,col,sheet)
          file.print one_cell_output(onecelltype,onecell,empty?(row,col,sheet))
        end
        file.print("\n")
      end # sheet not empty
    end
  end

  def one_cell_output(onecelltype,onecell,empty)
    str = ""
    if empty
      str += ''
    else
      case onecelltype
      when :string
        if onecell == ""
          str << ''
        else
          onecell.gsub!(/"/,'""')
          str << ('"'+onecell+'"')
        end
      when :float,:percentage
        if onecell == onecell.to_i
          str << onecell.to_i.to_s
        else
          str << onecell.to_s
        end
      when :formula
        if onecell.class == String
          if onecell == ""
            str << ''
          else
            onecell.gsub!(/"/,'""')
            str << '"'+onecell+'"'
          end
        elsif onecell.class == Float
          if onecell == onecell.to_i
            str << onecell.to_i.to_s
          else
            str << onecell.to_s
          end
        else
          raise "unhandled onecell-class "+onecell.class.to_s
        end
      when :date
        str << onecell.to_s
      when :time
        str << GenericSpreadsheet.integer_to_timestring(onecell)
      else
        raise "unhandled celltype "+onecelltype.to_s
      end
    end
    str
  end

  # converts an integer value to a time string like '02:05:06'
  def self.integer_to_timestring(content)
    h = (content/3600.0).floor
    content = content - h*3600
    m = (content/60.0).floor
    content = content - m*60
    s = content
    sprintf("%02d:%02d:%02d",h,m,s)
  end
end
