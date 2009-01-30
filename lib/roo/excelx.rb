
require 'rubygems'
require 'rexml/document'
require 'fileutils'
require 'zip/zipfilesystem'
require 'date'

class  String
  def end_with?(str)
    self[-str.length,str.length] == str
  end
end

class Excelx < GenericSpreadsheet
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
    @tmpdir = "oo_"+$$.to_s
    @tmpdir = File.join(ENV['ROO_TMP'], @tmpdir) if ENV['ROO_TMP'] 
    unless File.exists?(@tmpdir)
      FileUtils::mkdir(@tmpdir)
    end
    filename = open_from_uri(filename) if filename[0,7] == "http://"
    filename = unzip(filename) if packed and packed == :zip
    begin
      file_type_check(filename,'.xlsx','an Excel-xlsx')
      @cells_read = Hash.new
      @filename = filename
      unless File.file?(@filename)
        raise IOError, "file #{@filename} does not exist"
      end
      @@nr += 1
      @file_nr = @@nr
      extract_content(@filename)
      file = File.new(File.join(@tmpdir, @file_nr.to_s+"_roo_workbook.xml"))
      @workbook_doc = REXML::Document.new file
      file.close
      @shared_table = []
      if File.exist?(File.join(@tmpdir, @file_nr.to_s+'_roo_sharedStrings.xml'))
        file = File.new(File.join(@tmpdir, @file_nr.to_s+'_roo_sharedStrings.xml'))
        @sharedstring_doc = REXML::Document.new file
        file.close
        read_shared_strings(@sharedstring_doc)
      end
      @styles_table = []
      if File.exist?(File.join(@tmpdir, @file_nr.to_s+'_roo_styles.xml'))
        file = File.new(File.join(@tmpdir, @file_nr.to_s+'_roo_styles.xml'))
        @styles_doc = REXML::Document.new file
        file.close
        read_styles(@styles_doc)
      end
      @sheet_doc = []
      @sheet_files.each_with_index do |item, i|
        file = File.new(item)
        @sheet_doc[i] = REXML::Document.new file
        file.close
      end
    ensure
      #if ENV["roo_local"] != "thomas-p"
      FileUtils::rm_r(@tmpdir)
      #end
    end
    @default_sheet = nil
    # no need to set default_sheet if there is only one sheet in the document
    if self.sheets.size == 1
      @default_sheet = self.sheets.first
    end
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
    result = attribute2format(s)
    result
  end
  
  # returns an array of sheet names in the spreadsheet
  def sheets
    return_sheets = []
    @workbook_doc.each_element do |workbook|
      workbook.each_element do |el|
        if el.name == "sheets"
          el.each_element do |sheet|
            return_sheets << sheet.attributes['name']
          end
        end
      end
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

  # returns each formula in the selected sheet as an array of elements
  # [row, col, formula]
  def formulas(sheet=nil)
    theformulas = Array.new
    sheet = @default_sheet unless sheet
    read_cells(sheet) unless @cells_read[sheet]
    first_row(sheet).upto(last_row(sheet)) {|row|
      first_column(sheet).upto(last_column(sheet)) {|col|
        if formula?(row,col,sheet)
          f = [row, col, formula(row,col,sheet)]
          theformulas << f
        end
      }
    }
    theformulas
  end

  private

  # helper function to set the internal representation of cells
  def set_cell_values(sheet,x,y,i,v,vt,formula,tr,str_v,
      excelx_type=nil,
      excelx_value=nil,
      s_attribute=nil)
    key = [y,x+i]
    @cell_type[sheet] = {} unless @cell_type[sheet]
    @cell_type[sheet][key] = vt
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

  # splits a coordinate like "AA12" into the parts "AA" (String) and 12 (Fixnum)
  def split_coord(s)
    letter = ""
    number = 0
    i = 0
    while i<s.length and "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz".include?(s[i,1])
      letter += s[i,1]
      i+=1
    end
    while i<s.length and "0123456789".include?(s[i,1])
      number = number*10 + s[i,1].to_i
      i+=1
    end
    if letter=="" or number==0
      raise ArgumentError
    end
    return letter,number
  end

  def split_coordinate(str)
    letter,number = split_coord(str)
    x = GenericSpreadsheet.letter_to_number(letter)
    y = number
    return x,y
  end

  # read all cells in the selected sheet
  def format2type(format)
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
    @sheet_doc[n].each_element do |worksheet|
      worksheet.each_element do |elem|
        if elem.name == 'sheetData'
          elem.each_element do |sheetdata|
            if sheetdata.name == 'row'
              sheetdata.each_element do |row|
                if row.name == 'c'
                  if row.attributes['t'] == 's'
                    tmp_type = :shared
                  else
                    s_attribute = row.attributes['s']
                    format = attribute2format(s_attribute)
                    tmp_type = format2type(format)
                  end
                  formula = nil
                  row.each_element do |cell|
#                    puts "cell.name: #{cell.name}" if cell.text.include? "22606.5120"
#                    puts "cell.text: #{cell.text}" if cell.text.include? "22606.5120"
                    if cell.name == 'f'
                      formula = cell.text
                    end
                    if cell.name == 'v'
                      #puts "tmp_type: #{tmp_type}" if cell.text.include? "22606.5120"
                      #puts cell.name
                      if tmp_type == :time or tmp_type == :datetime #2008-07-26
                        #p cell.text
                       # p cell.text.to_f if cell.text.include? "22606.5120"
                        if cell.text.to_f >= 1.0 # 2008-07-26
                        #  puts ">= 1.0" if cell.text.include? "22606.5120"
                         # puts "cell.text.to_f: #{cell.text.to_f}" if cell.text.include? "22606.5120"
                          #puts "cell.text.to_f.floor: #{cell.text.to_f.floor}" if cell.text.include? "22606.5120"
                          if (cell.text.to_f - cell.text.to_f.floor).abs > 0.000001 #TODO: 
                           # puts "abs ist groesser"  if cell.text.include? "22606.5120"
                            # @cell[sheet][key] = DateTime.parse(tr.attributes['date-value'])
                            tmp_type = :datetime
                            
                          else
                            #puts ":date"
                            tmp_type = :date # 2008-07-26
                          end
                        else
                          #puts "<1.0"
                        end # 2008-07-26
                      end # 2008-07-26
                      excelx_type = [:numeric_or_formula,format]
                      excelx_value = cell.text
                      if tmp_type == :shared
                        vt = :string
                        str_v = @shared_table[cell.text.to_i]
                        excelx_type = :string
                      elsif tmp_type == :date
                        vt = :date
                        v = cell.text
                      elsif tmp_type == :time
                        vt = :time
                        v = cell.text
                      elsif tmp_type == :datetime
                        vt = :datetime
                        v = cell.text
                      elsif tmp_type == :formula
                        vt = :formula
                        v = cell.text.to_f #TODO: !!!!
                      else
                        vt = :float
                        v = cell.text
                      end
                      #puts "vt: #{vt}" if cell.text.include? "22606.5120"
                      x,y = split_coordinate(row.attributes['r'])
                      tr=nil #TODO: ???s
                      set_cell_values(sheet,x,y,0,v,vt,formula,tr,str_v,excelx_type,excelx_value,s_attribute)
                    end
                  end
                end
              end
            end
          end
        end
      end
    end
    sheet_found = true #TODO:
    if !sheet_found
      raise RangeError
    end
    @cells_read[sheet] = true
  end
  
  # Checks if the default_sheet exists. If not an RangeError exception is
  # raised
  def check_default_sheet
    sheet_found = false
    raise ArgumentError, "Error: default_sheet not set" if @default_sheet == nil
    @workbook_doc.each_element do |workbook|
      workbook.each_element do |el|
        if el.name == "sheets"
          el.each_element do |sheet|
            if @default_sheet == sheet.attributes['name']
              sheet_found = true
            end
          end
        end
      end
    end
    if ! sheet_found
      raise RangeError, "sheet '#{@default_sheet}' not found"
    end
  end

  # extracts all needed files from the zip file
  def process_zipfile(zipfilename, zip, path='')
    @sheet_files = []
    Zip::ZipFile.open(zipfilename) {|zf|
      zf.entries.each {|entry|
        #entry.extract
        if entry.to_s.end_with?('workbook.xml')
          open(@tmpdir+'/'+@file_nr.to_s+'_roo_workbook.xml','wb') {|f|
            f << zip.read(entry)
          }
        end
        if entry.to_s.end_with?('sharedStrings.xml')
          open(@tmpdir+'/'+@file_nr.to_s+'_roo_sharedStrings.xml','wb') {|f|
            f << zip.read(entry)
          }
        end
        if entry.to_s.end_with?('styles.xml')
          open(@tmpdir+'/'+@file_nr.to_s+'_roo_styles.xml','wb') {|f|
            f << zip.read(entry)
          }
        end
        if entry.to_s =~ /sheet([0-9]+).xml$/
          nr = $1
          open(@tmpdir+'/'+@file_nr.to_s+"_roo_sheet#{nr}",'wb') {|f|
            f << zip.read(entry)
          }
          @sheet_files[nr.to_i-1] = @tmpdir+'/'+@file_nr.to_s+"_roo_sheet#{nr}"
        end
      }
    }
    return
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
    doc.each_element do |sst|
      if sst.name == 'sst'
        sst.each_element do |si|
          if si.name == 'si'
            si.each_element do |elem|
              if elem.name == 't'
                @shared_table << elem.text
              end
            end
          end
        end
      end
    end
  end

  # read the styles elements of an excelx document
  def read_styles(doc)
    @numFmts = []
    @cellXfs = []
    doc.each_element do |e1|
      if e1.name == "styleSheet"
        e1.each_element do |e2|
          if e2.name == "numFmts"
            e2.each_element do |e3|
              if e3.name == 'numFmt'
                numFmtId = e3.attributes['numFmtId']
                formatCode = e3.attributes['formatCode']
                @numFmts << [numFmtId, formatCode]
              end
            end
          elsif e2.name == "cellXfs"
            e2.each_element do |e3|
              if e3.name == 'xf'
                numFmtId = e3.attributes['numFmtId'] 
                @cellXfs << [numFmtId]
              end
            end
          end
        end
      end
    end
  end

  # convert internal excelx attribute to a format
  def attribute2format(s)
    result = nil
    @numFmts.each {|nf|
      if nf.first == @cellXfs[s.to_i].first
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
