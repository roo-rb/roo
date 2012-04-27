# encoding: utf-8

require 'tmpdir'

# Base class for all other types of spreadsheets
class Roo::GenericSpreadsheet
  TEMP_PREFIX = "oo_"

  attr_reader :default_sheet, :headers

  # sets the line with attribute names (default: 1)
  attr_accessor :header_line

  protected

  def self.split_coordinate(str)
    letter,number = Roo::GenericSpreadsheet.split_coord(str)
    x = letter_to_number(letter)
    y = number
    return y, x
  end

  def self.split_coord(s)
    if s =~ /([a-zA-Z]+)([0-9]+)/
      letter = $1
      number = $2.to_i
    else
      raise ArgumentError
    end
    return letter, number
  end


  public

  # sets the working sheet in the document
  # 'sheet' can be a number (1 = first sheet) or the name of a sheet.
  def default_sheet=(sheet)
    validate_sheet!(sheet)
    @default_sheet = sheet
    @first_row[sheet] = @last_row[sheet] = @first_column[sheet] = @last_column[sheet] = nil
    @cells_read[sheet] = false
  end

  # first non-empty column as a letter
  def first_column_as_letter(sheet=nil)
    Roo::GenericSpreadsheet.number_to_letter(first_column(sheet))
  end

  # last non-empty column as a letter
  def last_column_as_letter(sheet=nil)
    Roo::GenericSpreadsheet.number_to_letter(last_column(sheet))
  end

  # returns the number of the first non-empty row
  def first_row(sheet=nil)
    sheet ||= @default_sheet
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
    sheet ||= @default_sheet
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
    sheet ||= @default_sheet
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
    sheet ||= @default_sheet
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
    sheet ||= @default_sheet
    result = "--- \n"
    return '' unless first_row # empty result if there is no first_row in a sheet

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
            result << "  value: #{Roo::GenericSpreadsheet.integer_to_timestring( self.cell(row,col,sheet))} \n"
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
    sheet ||= @default_sheet
    if filename
      File.open(filename,"w") do |file|
        write_csv_content(file,sheet)
      end
      return true
    else
      sio = StringIO.new
      write_csv_content(sio,sheet)
      sio.rewind
      return sio.read
    end
  end

  # returns a matrix object from the whole sheet or a rectangular area of a sheet
  def to_matrix(from_row=nil, from_column=nil, to_row=nil, to_column=nil,sheet=nil)
    require 'matrix'

    sheet ||= @default_sheet
    return Matrix.empty unless first_row

    Matrix.rows((from_row||first_row(sheet)).upto(to_row||last_row(sheet)).map do |row|
      (from_column||first_column(sheet)).upto(to_column||last_column(sheet)).map do |col|
        cell(row,col)
      end
    end)
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
    sheet ||= @default_sheet
    read_cells(sheet) unless @cells_read[sheet]
    first_column(sheet).upto(last_column(sheet)).map do |col|
      cell(rownumber,col,sheet)
    end
  end

  # returns all values in this column as an array
  # column numbers are 1,2,3,... like in the spreadsheet
  def column(columnnumber,sheet=nil)
    if columnnumber.class == String
      columnnumber = Roo::Excel.letter_to_number(columnnumber)
    end
    sheet ||= @default_sheet
    read_cells(sheet) unless @cells_read[sheet]
    first_row(sheet).upto(last_row(sheet)).map do |row|
      cell(row,columnnumber,sheet)
    end
  end

  # reopens and read a spreadsheet document
  def reload
    # von Abfrage der Klasse direkt auf .to_s == '..' umgestellt
    ds = @default_sheet
    if self.class.to_s == 'Google'
      initialize(@spreadsheetkey,@user,@password)
    else
      initialize(@filename)
    end
    self.default_sheet = ds
    #@first_row = @last_row = @first_column = @last_column = nil
  end

  # true if cell is empty
  def empty?(row, col, sheet=nil)
    sheet ||= @default_sheet
    read_cells(sheet) unless @cells_read[sheet] or self.class == Roo::Excel
    row,col = normalize(row,col)
    return true unless cell(row, col, sheet)
    return true if celltype(row, col, sheet) == :string && cell(row, col, sheet).empty?
    return true if row < first_row(sheet) || row > last_row(sheet) || col < first_column(sheet) || col > last_column(sheet)
    false
  end

  # returns information of the spreadsheet document and all sheets within
  # this document.
  def info
    result = "File: #{File.basename(@filename)}\n"+
      "Number of sheets: #{sheets.size}\n"+
      "Sheets: #{sheets.join(', ')}\n"
    n = 1
    sheets.each {|sheet|
      self.default_sheet = sheet
      result << "Sheet " + n.to_s + ":\n"
      unless first_row
        result << "  - empty -"
      else
        result << "  First row: #{first_row}\n"
        result << "  Last row: #{last_row}\n"
        result << "  First column: #{Roo::GenericSpreadsheet.number_to_letter(first_column)}\n"
        result << "  Last column: #{Roo::GenericSpreadsheet.number_to_letter(last_column)}"
      end
      result << "\n" if sheet != sheets.last
      n += 1
    }
    result
  end

  # returns an XML representation of all sheets of a spreadsheet file
  def to_xml
    builder = Nokogiri::XML::Builder.new do |xml|
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
    end
    return builder.to_xml
  end

  # when a method like spreadsheet.a42 is called
  # convert it to a call of spreadsheet.cell('a',42)
  def method_missing(m, *args)
    # #aa42 => #cell('aa',42)
    # #aa42('Sheet1')  => #cell('aa',42,'Sheet1')
    if m =~ /^([a-z]+)(\d)$/
      col = Roo::GenericSpreadsheet.letter_to_number($1)
      row = $2.to_i
      if args.size > 0
        return cell(row,col,args[0])
      else
        return cell(row,col)
      end
    else
      super
    end
  end

=begin
#TODO: hier entfernen
  # returns each formula in the selected sheet as an array of elements
  # [row, col, formula]
  def formulas(sheet=nil)
    theformulas = Array.new
    sheet ||= @default_sheet
    read_cells(sheet) unless @cells_read[sheet]
    return theformulas unless first_row(sheet) # if there is no first row then
    # there can't be formulas
    first_row(sheet).upto(last_row(sheet)) {|row|
      first_column(sheet).upto(last_column(sheet)) {|col|
        if formula?(row,col,sheet)
          theformulas << [row, col, formula(row,col,sheet)]
        end
      }
    }
    theformulas
  end
=end



    # FestivalBobcats fork changes begin here



    # access different worksheets by calling spreadsheet.sheet(1)
    # or spreadsheet.sheet('SHEETNAME')
    def sheet(index,name=false)
      @default_sheet = String === index ? index : self.sheets[index]
      name ? [@default_sheet,self] : self
    end

    # iterate through all worksheets of a document
    def each_with_pagename
      self.sheets.each do |s|
        yield sheet(s,true)
      end
    end

    # by passing in headers as options, this method returns
    # specific columns from your header assignment
    # for example:
    # xls.sheet('New Prices').parse(:upc => 'UPC', :price => 'Price') would return:
    # [{:upc => 123456789012, :price => 35.42},..]

    # the queries are matched with regex, so regex options can be passed in
    # such as :price => '^(Cost|Price)'
    # case insensitive by default


    # by using the :header_search option, you can query for headers
    # and return a hash of every row with the keys set to the header result
    # for example:
    # xls.sheet('New Prices').parse(:header_search => ['UPC*SKU','^Price*\sCost\s'])

    # that example searches for a column titled either UPC or SKU and another
    # column titled either Price or Cost (regex characters allowed)
    # * is the wildcard character

    # you can also pass in a :clean => true option to strip the sheet of
    # odd unicode characters and white spaces around columns

    def each(options={})
      if options.empty?
        1.upto(last_row) do |line|
          yield row(line)
        end
      else
        if options[:clean]
          options.delete(:clean)
          @cleaned ||= {}
          @cleaned[@default_sheet] || clean_sheet
        end

        if options[:header_search]
          @headers = nil
          @header_line = row_with(options[:header_search])
        elsif [:first_row,true].include?(options[:headers])
          @headers = []
          row(first_row).each_with_index {|x,i| @headers << [x,i + 1]}
        else
          set_headers(options)
        end

        headers = @headers ||
          Hash[(first_column..last_column).map do |col|
            [cell(@header_line,col), col]
          end]

        @header_line.upto(last_row) do |line|
          yield(Hash[headers.map {|k,v| [k,cell(line,v)]}])
        end
      end
    end

    def parse(options={})
      ary = []
      if block_given?
        each(options) {|row| ary << yield(row)}
      else
        each(options) {|row| ary << row}
      end
      ary
    end

    def row_with(query,return_headers=false)
      query.map! {|x| Array(x.split('*'))}
      line_no = 0
      each do |row|
        line_no += 1
        # makes sure headers is the first part of wildcard search for priority
        # ex. if UPC and SKU exist for UPC*SKU search, UPC takes the cake
        headers = query.map do |q|
          q.map {|i| row.grep(/#{i}/i)[0]}.compact[0]
        end.compact

        if headers.length == query.length
          @header_line = line_no
          return_headers ? (return headers) : (return line_no)
        end
        raise "Couldn't find header row." if line_no > 100
      end
    end

    # this method lets you find the worksheet with the most data
    def longest_sheet
      sheet(@workbook.worksheets.inject {|m,o| o.row_count > m.row_count ? o : m}.name)
    end

  protected

  def file_type_check(filename, ext, name, packed=nil)
    new_expression = {
      '.ods' => 'Roo::Openoffice.new',
      '.xls' => 'Roo::Excel.new',
      '.xlsx' => 'Roo::Excelx.new',
      '.csv' => 'Roo::Csv.new',
    }
    if packed == :zip
	    # lalala.ods.zip => lalala.ods
	    # hier wird KEIN unzip gemacht, sondern nur der Name der Datei
	    # getestet, falls es eine gepackte Datei ist.
	    filename = File.basename(filename,File.extname(filename))
    end
    case ext
    when '.ods', '.xls', '.xlsx', '.csv'
      correct_class = "use #{new_expression[ext]} to handle #{ext} spreadsheet files. This has #{File.extname(filename).downcase}"
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

  # see: key_to_num
  def key_to_string(arr)
    "#{arr[0]},#{arr[1]}"
  end

  private

  def make_tmpdir(tmp_root = nil)
    Dir.mktmpdir(TEMP_PREFIX, tmp_root || ENV['ROO_TMP']) do |tmpdir|
      yield tmpdir
    end
  end

  def clean_sheet
    read_cells(@default_sheet) unless @cells_read[@default_sheet]
    @cell = Hash[
              @cell.map do |k,v|
                [k,Hash[v.map {|k,v| [k,sanitize_value(v)]}]]
              end]
    @cleaned[@default_sheet] = true
  end

  def sanitize_value(v)
    String === v ? v.strip.unpack('U*').select {|b| b < 127}.pack('U*') : v
  end

  def set_headers(hash={})
    # try to find header row with all values or give an error
    # then create new hash by indexing strings and keeping integers for header array
    @headers = row_with(hash.values,true)
    @headers = Hash[hash.keys.zip(@headers.map {|x| header_index(x)})]
  end

  def header_index(query)
    row(@header_line).index(query) + first_column
  end

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
      col = Roo::GenericSpreadsheet.letter_to_number(col)
    end
    return row,col
  end

  def open_from_uri(uri, tmpdir)
    require 'open-uri'
    response = ''
    begin
      open(uri, "User-Agent" => "Ruby/#{RUBY_VERSION}") { |net|
        response = net.read
        tempfilename = File.join(tmpdir, File.basename(uri))
        File.open(tempfilename,"wb") do |file|
          file.write(response)
        end
      }
    rescue OpenURI::HTTPError
      raise "could not open #{uri}"
    end
    File.join(tmpdir, File.basename(uri))
  end

  def open_from_stream(stream, tmpdir)
    tempfilename = File.join(tmpdir, "spreadsheet")
    File.open(tempfilename,"wb") do |file|
      file.write(stream[7..-1])
    end
    File.join(tmpdir, "spreadsheet")
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

  def unzip(filename, tmpdir)
    ret = nil
    Zip::ZipFile.open(filename) do |zip|
      ret = process_zipfile_packed(zip, tmpdir)
    end
    ret
  end

  # check if default_sheet was set and exists in sheets-array
  def validate_sheet!(sheet)
    case sheet
    when nil
      raise ArgumentError, "Error: sheet 'nil' not valid"
    when Fixnum
      self.sheets.fetch(sheet-1) do
        raise RangeError, "sheet index #{sheet} not found"
      end
    when String
      if !sheets.include? sheet
        raise RangeError, "sheet '#{sheet}' not found"
      end
    else
      raise TypeError, "not a valid sheet type: #{sheet.inspect}"
    end
  end

  def process_zipfile_packed(zip, tmpdir, path='')
    ret=nil
    if zip.file.file? path
      # extract and return filename
      File.open(File.join(tmpdir, path),"wb") do |file|
        file.write(zip.read(path))
      end
      return File.join(tmpdir, path)
    else
      path += '/' unless path.empty?
      zip.dir.foreach(path) do |filename|
        ret = process_zipfile_packed(zip, tmpdir, path + filename)
      end
    end
    ret
  end

  # Write all cells to the csv file. File can be a filename or nil. If the this
  # parameter is nil the output goes to STDOUT
  def write_csv_content(file=nil,sheet=nil)
    file ||= STDOUT
    if first_row(sheet) # sheet is not empty
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

  # The content of a cell in the csv output
  def one_cell_output(onecelltype, onecell, empty)
    str = ""
    if empty
      str += ''
    else
      case onecelltype
      when :string
        unless onecell.empty?
          one = onecell.gsub(/"/,'""')
          str << ('"'+one+'"')
        end
      when :float, :percentage
        if onecell == onecell.to_i
          str << onecell.to_i.to_s
        else
          str << onecell.to_s
        end
      when :formula
        if onecell.class == String
          unless onecell.empty?
            one = onecell.gsub(/"/,'""')
            str << '"'+one+'"'
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
        str << Roo::GenericSpreadsheet.integer_to_timestring(onecell)
      when :datetime
        str << onecell.to_s
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
