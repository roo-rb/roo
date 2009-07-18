#damit keine falschen Vermutungen aufkommen: Ich habe religioes rein gar nichts
# mit diesem Bibelbund zu tun, aber die hatten eine ziemlich grosse
# Spreadsheet-Datei mit ca. 3500 Zeilen oeffentlich im Netz, die sich ganz gut
# zum Testen eignete.
#
#--
# these test cases were developed to run under Linux OS, some commands
# (like 'diff') must be changed (or commented out ;-)) if you want to run
# the tests under another OS
#


#TODO
# Look at formulas in excel - does not work with date/time


# Dump warnings that come from the test to open files
# with the wrong spreadsheet class
STDERR.reopen "/dev/null","w"

TESTDIR =  File.dirname(__FILE__) 
require TESTDIR + '/test_helper.rb'
#require 'soap/rpc/driver'
require 'fileutils'
require 'timeout'
require 'logger'
$log = Logger.new(File.join(ENV['HOME'],"roo.log"))
$log.level = Logger::WARN
#$log.level = Logger::DEBUG

DISPLAY_LOG = false
DB_LOG = false

if DB_LOG
  require 'activerecord'
end

include FileUtils

if DB_LOG
  def activerecord_connect
    ActiveRecord::Base.establish_connection(:adapter => "mysql",
      :database => "test_runs",
      :host => "localhost",
      :username => "root",
      :socket => "/var/run/mysqld/mysqld.sock")
  end

  class Testrun < ActiveRecord::Base
  end
end

class Test::Unit::TestCase
  def key_of(spreadsheetname)
    begin
      
      return {
        'formula' => 'rt4Pw1WmjxFtyfrqqy94wPw',
        "write.me" => 'r6m7HFlUOwst0RTUTuhQ0Ow',
        'numbers1' => "rYraCzjxTtkxw1NxHJgDU8Q",
        'borders' => "r_nLYMft6uWg_PT9Rc2urXw",
        'simple_spreadsheet' => "r3aMMCBCA153TmU_wyIaxfw",
        'testnichtvorhandenBibelbund.ods' => "invalidkeyforanyspreadsheet", # !!! intentionally false key
        "only_one_sheet" => "rqRtkcPJ97nhQ0m9ksDw2rA",
        'time-test' => 'r2XfDBJMrLPjmuLrPQQrEYw',
        'datetime' => "r2kQpXWr6xOSUpw9MyXavYg",
        'whitespace' => "rZyQaoFebVGeHKzjG6e9gRQ"
      }[spreadsheetname]
        # 'numbers1' => "o10837434939102457526.4784396906364855777",
        # 'borders' => "o10837434939102457526.664868920231926255",
        # 'simple_spreadsheet' => "ptu6bbahNZpYe-L1vEBmgGA",
        # 'testnichtvorhandenBibelbund.ods' => "invalidkeyforanyspreadsheet", # !!! intentionally false key
        # "only_one_sheet" => "o10837434939102457526.762705759906130135",
        # "write.me" => 'ptu6bbahNZpY0N0RrxQbWdw&hl',
        # 'formula' => 'o10837434939102457526.3022866619437760118',
        # 'time-test' => 'ptu6bbahNZpYBMhk01UfXSg',
        # 'datetime' => "ptu6bbahNZpYQEtZwzL_dZQ",
    rescue
      raise "unknown spreadsheetname: #{spreadsheetname}"
    end
  end

  if DB_LOG
    if ! (defined?(@connected) and @connected)
      activerecord_connect
    else
      @connected = true
    end
  end
  alias unlogged_run run
  def run(result, &block)
    t1 = Time.now
    #RAILS_DEFAULT_LOGGER.debug "RUNNING #{self.class} #{@method_name} \t#{Time.now.to_s}"
    if DISPLAY_LOG
      print "RUNNING #{self.class} #{@method_name} \t#{Time.now.to_s}"
      STDOUT.flush
    end
    unlogged_run result, &block
    t2 = Time.now
    if DISPLAY_LOG
      puts "\t#{t2-t1} seconds"
    end
    if DB_LOG
      domain = Testrun.create(
        :class_name => self.class.to_s,
        :test_name => @method_name,
        :start => t1,
        :duration => t2-t1
      )
    end
  end
end

class File
  def File.delete_if_exist(filename)
    if File.exist?(filename)
      File.delete(filename)
    end
  end
end

# :nodoc
class Fixnum
  def minutes
    self * 60
  end
end

class TestRoo < Test::Unit::TestCase

  OPENOFFICE   = true  	# do Openoffice-Spreadsheet Tests?
  EXCEL        = true	  # do Excel Tests?
  GOOGLE       = false 	# do Google-Spreadsheet Tests?
  EXCELX       = true  	# do Excel-X Tests? (.xlsx-files)

  ONLINE = true
  LONG_RUN = false
  GLOBAL_TIMEOUT = 48.minutes #*60 # 2*12*60 # seconds

  def setup
    #if DISPLAY_LOG
    #  puts " GLOBAL_TIMEOUT = #{GLOBAL_TIMEOUT}"
    #end
  end

  def test_internal_minutes
    assert_equal 42*60, 42.minutes
  end
  
  # call a block of code for each spreadsheet type 
  # and yield a reference to the roo object
  def with_each_spreadsheet(options)
     options[:format] ||= [:excel, :excelx, :openoffice, :google]
     options[:format] = [options[:format]] if options[:format].class == Symbol
     yield Roo::Spreadsheet.open(File.join(TESTDIR, options[:name] + '.xls')) if EXCEL && options[:format].include?(:excel)
     yield Roo::Spreadsheet.open(File.join(TESTDIR, options[:name] + '.xlsx')) if EXCELX && options[:format].include?(:excelx)
     yield Roo::Spreadsheet.open(File.join(TESTDIR, options[:name] + '.ods')) if OPENOFFICE && options[:format].include?(:openoffice)
     yield Roo::Spreadsheet.open(key_of(options[:name]) || options[:name]) if GOOGLE && options[:format].include?(:google)
  end

  # Using Date.strptime so check that it's using the method
  # with the value set in date_format
  def test_date
    with_each_spreadsheet(:name=>'numbers1', :format=>:google) do |oo|
      # should default to  DDMMYYYY
      assert oo.date?("21/11/1962") == true
      assert oo.date?("11/21/1962") == false
      oo.date_format = '%m/%d/%Y'
      assert oo.date?("21/11/1962") == false
      assert oo.date?("11/21/1962") == true
      oo.date_format = '%Y-%m-%d'
      assert oo.date?("1962-11-21") == true
      assert oo.date?("1962-21-11") == false
    end
  end

  def test_classes
    if OPENOFFICE
      oo = Openoffice.new(File.join(TESTDIR,"numbers1.ods"))
      assert_kind_of Openoffice, oo
    end
    if EXCEL
      oo = Excel.new(File.join(TESTDIR,"numbers1.xls"))
      assert_kind_of Excel, oo
    end
    if GOOGLE
      oo = Google.new(key_of("numbers1"))
      assert_kind_of Google, oo
    end
    if EXCELX
      oo = Excelx.new(File.join(TESTDIR,"numbers1.xlsx"))
      assert_kind_of Excelx, oo
    end
  end

  def test_letters
    assert_equal 1, GenericSpreadsheet.letter_to_number('A')
    assert_equal 1, GenericSpreadsheet.letter_to_number('a')
    assert_equal 2, GenericSpreadsheet.letter_to_number('B')
    assert_equal 26, GenericSpreadsheet.letter_to_number('Z')
    assert_equal 27, GenericSpreadsheet.letter_to_number('AA')
    assert_equal 27, GenericSpreadsheet.letter_to_number('aA')
    assert_equal 27, GenericSpreadsheet.letter_to_number('Aa')
    assert_equal 27, GenericSpreadsheet.letter_to_number('aa')
  end

  def test_sheets
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      assert_equal ["Tabelle1","Name of Sheet 2","Sheet3","Sheet4","Sheet5"], oo.sheets
      assert_raise(RangeError) { oo.default_sheet = "no_sheet" }
      assert_raise(TypeError)  { oo.default_sheet = [1,2,3] }
      oo.sheets.each { |sh|
        oo.default_sheet = sh
        assert_equal sh, oo.default_sheet
      }
    end  
  end
  
  def test_cells
    with_each_spreadsheet(:name=>'numbers1') do |oo| 
      assert_equal 1, oo.cell(1,1)
      assert_equal 2, oo.cell(1,2)
      assert_equal 3, oo.cell(1,3)
      assert_equal 4, oo.cell(1,4)
      assert_equal 5, oo.cell(2,1)
      assert_equal 6, oo.cell(2,2)
      assert_equal 7, oo.cell(2,3)
      assert_equal 8, oo.cell(2,4)
      assert_equal 9, oo.cell(2,5)
      assert_equal "test", oo.cell(2,6)
      assert_equal :string, oo.celltype(2,6)
      assert_equal 11, oo.cell(2,7)
      assert_equal :float, oo.celltype(2,7)
      assert_equal 10, oo.cell(4,1)
      assert_equal 11, oo.cell(4,2)
      assert_equal 12, oo.cell(4,3)
      assert_equal 13, oo.cell(4,4)
      assert_equal 14, oo.cell(4,5)
      assert_equal 10, oo.cell(4,'A')
      assert_equal 11, oo.cell(4,'B')
      assert_equal 12, oo.cell(4,'C')
      assert_equal 13, oo.cell(4,'D')
      assert_equal 14, oo.cell(4,'E')
      assert_equal :date, oo.celltype(5,1)
      assert_equal Date.new(1961,11,21), oo.cell(5,1)
      assert_equal "1961-11-21", oo.cell(5,1).to_s
    end
  end
  
  def test_celltype
    with_each_spreadsheet(:name=>'numbers1') do |oo| 
      assert_equal :string, oo.celltype(2,6)
    end  
  end

  def test_cell_address
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      assert_equal "tata", oo.cell(6,1)
      assert_equal "tata", oo.cell(6,'A')
      assert_equal "tata", oo.cell('A',6)
      assert_equal "tata", oo.cell(6,'a')
      assert_equal "tata", oo.cell('a',6)
      assert_raise(ArgumentError) { assert_equal "tata", oo.cell('a','f') }
      assert_raise(ArgumentError) { assert_equal "tata", oo.cell('f','a') }
      assert_equal "thisisc8", oo.cell(8,3)
      assert_equal "thisisc8", oo.cell(8,'C')
      assert_equal "thisisc8", oo.cell('C',8)
      assert_equal "thisisc8", oo.cell(8,'c')
      assert_equal "thisisc8", oo.cell('c',8)
      assert_equal "thisisd9", oo.cell('d',9)
      assert_equal "thisisa11", oo.cell('a',11)
    end
  end

  def test_office_version
    with_each_spreadsheet(:name=>'numbers1', :format=>:openoffice) do |oo|
      assert_equal "1.0", oo.officeversion
    end  
  end

  #TODO: inkonsequente Lieferung Fixnum/Float
  def test_rows
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      assert_equal 41, oo.cell('a',12)
      assert_equal 42, oo.cell('b',12)
      assert_equal 43, oo.cell('c',12)
      assert_equal 44, oo.cell('d',12)
      assert_equal 45, oo.cell('e',12)
      assert_equal [41.0,42.0,43.0,44.0,45.0, nil, nil], oo.row(12)
      assert_equal "einundvierzig", oo.cell('a',16)
      assert_equal "zweiundvierzig", oo.cell('b',16)
      assert_equal "dreiundvierzig", oo.cell('c',16)
      assert_equal "vierundvierzig", oo.cell('d',16)
      assert_equal "fuenfundvierzig", oo.cell('e',16)
      assert_equal ["einundvierzig", "zweiundvierzig", "dreiundvierzig", "vierundvierzig", "fuenfundvierzig", nil, nil], oo.row(16)
    end
  end

  def test_last_row
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      assert_equal 18, oo.last_row
    end
  end

  def test_last_column
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      assert_equal 7, oo.last_column
    end
  end

  def test_last_column_as_letter
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      assert_equal 'G', oo.last_column_as_letter
    end
  end

  def test_first_row
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      assert_equal 1, oo.first_row
    end
  end

  def test_first_column
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      assert_equal 1, oo.first_column
    end
  end

  def test_first_column_as_letter
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      assert_equal 'A', oo.first_column_as_letter
    end
  end

  def test_sheetname
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      oo.default_sheet = "Name of Sheet 2"
      assert_equal 'I am sheet 2', oo.cell('C',5)
      assert_raise(RangeError) { oo.default_sheet = "non existing sheet name" }
      assert_raise(RangeError) { oo.default_sheet = "non existing sheet name" }
      assert_raise(RangeError) { dummy = oo.cell('C',5,"non existing sheet name")}
      assert_raise(RangeError) { dummy = oo.celltype('C',5,"non existing sheet name")}
      assert_raise(RangeError) { dummy = oo.empty?('C',5,"non existing sheet name")}
      if oo.class == Excel
        assert_raise(RuntimeError) { dummy = oo.formula?('C',5,"non existing sheet name")}
        assert_raise(RuntimeError) { dummy = oo.formula('C',5,"non existing sheet name")}
      else  
        assert_raise(RangeError) { dummy = oo.formula?('C',5,"non existing sheet name")}
        assert_raise(RangeError) { dummy = oo.formula('C',5,"non existing sheet name")}
        assert_raise(RangeError) { dummy = oo.set('C',5,42,"non existing sheet name")} unless oo.class == Google
        assert_raise(RangeError) { dummy = oo.formulas("non existing sheet name")} 
      end
      assert_raise(RangeError) { dummy = oo.to_yaml({},1,1,1,1,"non existing sheet name")}
    end
  end  

  def test_boundaries
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      oo.default_sheet = "Name of Sheet 2"
      assert_equal 2, oo.first_column
      assert_equal 'B', oo.first_column_as_letter
      assert_equal 5, oo.first_row
      assert_equal 'E', oo.last_column_as_letter
      assert_equal 14, oo.last_row
    end
  end

  def test_multiple_letters
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      oo.default_sheet = "Sheet3"
      assert_equal "i am AA", oo.cell('AA',1)
      assert_equal "i am AB", oo.cell('AB',1)
      assert_equal "i am BA", oo.cell('BA',1)
      assert_equal 'BA', oo.last_column_as_letter
      assert_equal "i am BA", oo.cell(1,'BA')
    end
  end

  def test_argument_error
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      assert_nothing_raised(ArgumentError) {  oo.default_sheet = "Tabelle1" }
    end
  end

  def test_empty_eh
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      assert oo.empty?('a',14)
      assert ! oo.empty?('a',15)
      assert oo.empty?('a',20)
    end
  end

  def test_reload
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      assert_equal 1, oo.cell(1,1)
      oo.reload
      assert_equal 1, oo.cell(1,1)
    end
  end

  def test_bug_contiguous_cells
    with_each_spreadsheet(:name=>'numbers1', :format=>:openoffice) do |oo|
      oo.default_sheet = "Sheet4"
      assert_equal Date.new(2007,06,16), oo.cell('a',1)
      assert_equal 10, oo.cell('b',1)
      assert_equal 10, oo.cell('c',1)
      assert_equal 10, oo.cell('d',1)
      assert_equal 10, oo.cell('e',1)
    end
  end

  def test_bug_italo_ve
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      oo.default_sheet = "Sheet5"
      assert_equal 1, oo.cell('A',1)
      assert_equal 5, oo.cell('b',1)
      assert_equal 5, oo.cell('c',1)
      assert_equal 2, oo.cell('a',2)
      assert_equal 3, oo.cell('a',3)
    end
  end

  def test_italo_table
    with_each_spreadsheet(:name=>'simple_spreadsheet_from_italo', :format=>[:openoffice, :excel]) do |oo|
      assert_equal  '1', oo.cell('A',1)
      assert_equal  '1', oo.cell('B',1)
      assert_equal  '1', oo.cell('C',1)
      assert_equal  1, oo.cell('A',2).to_i
      assert_equal  2, oo.cell('B',2).to_i
      assert_equal  1, oo.cell('C',2).to_i
      assert_equal  1, oo.cell('A',3)
      assert_equal  3, oo.cell('B',3)
      assert_equal  1, oo.cell('C',3)
      assert_equal  'A', oo.cell('A',4)
      assert_equal  'A', oo.cell('B',4)
      assert_equal  'A', oo.cell('C',4)
      assert_equal  0.01, oo.cell('A',5)
      assert_equal  0.01, oo.cell('B',5)
      assert_equal  0.01, oo.cell('C',5)
      assert_equal 0.03, oo.cell('a',5)+oo.cell('b',5)+oo.cell('c',5)

      # Cells values in row 1:
      assert_equal "1:string", oo.cell(1, 1)+":"+oo.celltype(1, 1).to_s
      assert_equal "1:string",oo.cell(1, 2)+":"+oo.celltype(1, 2).to_s
      assert_equal "1:string",oo.cell(1, 3)+":"+oo.celltype(1, 3).to_s

      # Cells values in row 2:
      assert_equal "1:string",oo.cell(2, 1)+":"+oo.celltype(2, 1).to_s
      assert_equal "2:string",oo.cell(2, 2)+":"+oo.celltype(2, 2).to_s
      assert_equal "1:string",oo.cell(2, 3)+":"+oo.celltype(2, 3).to_s

      # Cells values in row 3:
      assert_equal "1.0:float",oo.cell(3, 1).to_s+":"+oo.celltype(3, 1).to_s
      assert_equal "3.0:float",oo.cell(3, 2).to_s+":"+oo.celltype(3, 2).to_s
      assert_equal "1.0:float",oo.cell(3, 3).to_s+":"+oo.celltype(3, 3).to_s

      # Cells values in row 4:
      assert_equal "A:string",oo.cell(4, 1)+":"+oo.celltype(4, 1).to_s
      assert_equal "A:string",oo.cell(4, 2)+":"+oo.celltype(4, 2).to_s
      assert_equal "A:string",oo.cell(4, 3)+":"+oo.celltype(4, 3).to_s

      # Cells values in row 5:
      if oo.class == Openoffice
        assert_equal "0.01:percentage",oo.cell(5, 1).to_s+":"+oo.celltype(5, 1).to_s
        assert_equal "0.01:percentage",oo.cell(5, 2).to_s+":"+oo.celltype(5, 2).to_s
        assert_equal "0.01:percentage",oo.cell(5, 3).to_s+":"+oo.celltype(5, 3).to_s
      else
        assert_equal "0.01:float",oo.cell(5, 1).to_s+":"+oo.celltype(5, 1).to_s
        assert_equal "0.01:float",oo.cell(5, 2).to_s+":"+oo.celltype(5, 2).to_s
        assert_equal "0.01:float",oo.cell(5, 3).to_s+":"+oo.celltype(5, 3).to_s
      end
    end
  end

  def test_formula_openoffice
    with_each_spreadsheet(:name=>'formula', :format=>:openoffice) do |oo|
      assert_equal 1, oo.cell('A',1)
      assert_equal 2, oo.cell('A',2)
      assert_equal 3, oo.cell('A',3)
      assert_equal 4, oo.cell('A',4)
      assert_equal 5, oo.cell('A',5)
      assert_equal 6, oo.cell('A',6)
      assert_equal 21, oo.cell('A',7)
      assert_equal :formula, oo.celltype('A',7)
      assert_equal "=[Sheet2.A1]", oo.formula('C',7)
      assert_nil oo.formula('A',6)
      assert_equal [[7, 1, "=SUM([.A1:.A6])"],
        [7, 2, "=SUM([.$A$1:.B6])"],
        [7, 3, "=[Sheet2.A1]"],
        [8, 2, "=SUM([.$A$1:.B7])"],
      ], oo.formulas(oo.sheets.first)

      # setting a cell
      oo.set('A',15, 41)
      assert_equal 41, oo.cell('A',15)
      oo.set('A',16, "41")
      assert_equal "41", oo.cell('A',16)
      oo.set('A',17, 42.5)
      assert_equal 42.5, oo.cell('A',17)
    end
  end

  def test_formula_google
    with_each_spreadsheet(:name=>'formula', :format=>:google) do |oo|
      assert_equal 1, oo.cell('A',1)
      assert_equal 2, oo.cell('A',2)
      assert_equal 3, oo.cell('A',3)
      assert_equal 4, oo.cell('A',4)
      assert_equal 5, oo.cell('A',5)
      assert_equal 6, oo.cell('A',6)
      # assert_equal 21, oo.cell('A',7)
      assert_equal 21.0, oo.cell('A',7) #TODO: better solution Fixnum/Float
      assert_equal :formula, oo.celltype('A',7)
      # assert_equal "=[Sheet2.A1]", oo.formula('C',7)
      # !!! different from formulas in Openoffice
      #was: assert_equal "=sheet2!R[-6]C[-2]", oo.formula('C',7)
      # has Google changed their format of formulas/references to other sheets?
      assert_equal "=Sheet2!R[-6]C[-2]", oo.formula('C',7)
      assert_nil oo.formula('A',6)
      # assert_equal [[7, 1, "=SUM([.A1:.A6])"],
      #   [7, 2, "=SUM([.$A$1:.B6])"],
      #   [7, 3, "=[Sheet2.A1]"],
      #   [8, 2, "=SUM([.$A$1:.B7])"],
      # ], oo.formulas(oo.sheets.first)
      # different format than in openoffice spreadsheets:
      #was:
      # assert_equal [[7, 1, "=SUM(R[-6]C[0]:R[-1]C[0])"],
      #  [7, 2, "=SUM(R1C1:R[-1]C[0])"],
      #  [7, 3, "=sheet2!R[-6]C[-2]"],
      #  [8, 2, "=SUM(R1C1:R[-1]C[0])"]],
      #  oo.formulas(oo.sheets.first)
      assert_equal [[7, 1, "=SUM(R[-6]C:R[-1]C)"],
        [7, 2, "=SUM(R1C1:R[-1]C)"],
        [7, 3, "=Sheet2!R[-6]C[-2]"],
        [8, 2, "=SUM(R1C1:R[-1]C)"]],
        oo.formulas(oo.sheets.first)
    end
  end

  def test_formula_excelx
    with_each_spreadsheet(:name=>'formula', :format=>:excelx) do |oo|
      assert_equal 1, oo.cell('A',1)
      assert_equal 2, oo.cell('A',2)
      assert_equal 3, oo.cell('A',3)
      assert_equal 4, oo.cell('A',4)
      assert_equal 5, oo.cell('A',5)
      assert_equal 6, oo.cell('A',6)
      assert_equal 21, oo.cell('A',7)
      assert_equal :formula, oo.celltype('A',7)
      #steht nicht in Datei, oder?
      #nein, diesen Bezug habe ich nur in der Openoffice-Datei
      #assert_equal "=[Sheet2.A1]", oo.formula('C',7)
      assert_nil oo.formula('A',6)
      # assert_equal [[7, 1, "=SUM([.A1:.A6])"],
      #  [7, 2, "=SUM([.$A$1:.B6])"],
      #[7, 3, "=[Sheet2.A1]"],
      #[8, 2, "=SUM([.$A$1:.B7])"],
      #], oo.formulas(oo.sheets.first)
      assert_equal [[7, 1, 'SUM(A1:A6)'],
        [7, 2, 'SUM($A$1:B6)'],
        # [7, 3, "=[Sheet2.A1]"],
        # [8, 2, "=SUM([.$A$1:.B7])"],
      ], oo.formulas(oo.sheets.first)

      # setting a cell
      oo.set('A',15, 41)
      assert_equal 41, oo.cell('A',15)
      oo.set('A',16, "41")
      assert_equal "41", oo.cell('A',16)
      oo.set('A',17, 42.5)
      assert_equal 42.5, oo.cell('A',17)
    end
  end
  
  # Excel can only read the cell's value
  def test_formula_excel
    with_each_spreadsheet(:name=>'formula', :format=>:excel) do |oo|
      assert_equal 21, oo.cell('A',7)
      assert_equal 21, oo.cell('B',7)
    end
  end

  
  def test_borders_sheets
    with_each_spreadsheet(:name=>'borders') do |oo|
      oo.default_sheet = oo.sheets[1]
      assert_equal 6, oo.first_row
      assert_equal 11, oo.last_row
      assert_equal 4, oo.first_column
      assert_equal 8, oo.last_column

      oo.default_sheet = oo.sheets.first
      assert_equal 5, oo.first_row
      assert_equal 10, oo.last_row
      assert_equal 3, oo.first_column
      assert_equal 7, oo.last_column

      oo.default_sheet = oo.sheets[2]
      assert_equal 7, oo.first_row
      assert_equal 12, oo.last_row
      assert_equal 5, oo.first_column
      assert_equal 9, oo.last_column
    end
  end

  def yaml_entry(row,col,type,value)
    "cell_#{row}_#{col}: \n  row: #{row} \n  col: #{col} \n  celltype: #{type} \n  value: #{value} \n"
  end

  def test_to_yaml
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      assert_equal "--- \n"+yaml_entry(5,1,"date","1961-11-21"), oo.to_yaml({}, 5,1,5,1)
      assert_equal "--- \n"+yaml_entry(8,3,"string","thisisc8"), oo.to_yaml({}, 8,3,8,3)
      assert_equal "--- \n"+yaml_entry(12,3,"float",43.0), oo.to_yaml({}, 12,3,12,3)
      assert_equal \
        "--- \n"+yaml_entry(12,3,"float",43.0) +
        yaml_entry(12,4,"float",44.0) +
        yaml_entry(12,5,"float",45.0), oo.to_yaml({}, 12,3,12)
      assert_equal \
        "--- \n"+yaml_entry(12,3,"float",43.0)+
        yaml_entry(12,4,"float",44.0)+
        yaml_entry(12,5,"float",45.0)+
        yaml_entry(15,3,"float",43.0)+
        yaml_entry(15,4,"float",44.0)+
        yaml_entry(15,5,"float",45.0)+
        yaml_entry(16,3,"string","dreiundvierzig")+
        yaml_entry(16,4,"string","vierundvierzig")+
        yaml_entry(16,5,"string","fuenfundvierzig"), oo.to_yaml({}, 12,3)
    end
  end

  def test_only_one_sheet
    with_each_spreadsheet(:name=>'only_one_sheet') do |oo|
      assert_equal 42, oo.cell('B',4)
      assert_equal 43, oo.cell('C',4)
      assert_equal 44, oo.cell('D',4)
      oo.default_sheet = oo.sheets.first
      assert_equal 42, oo.cell('B',4)
      assert_equal 43, oo.cell('C',4)
      assert_equal 44, oo.cell('D',4)
    end
  end

  def test_excel_open_from_uri_and_zipped
    if EXCEL
      if ONLINE
        begin
          url = 'http://stiny-leonhard.de/bode-v1.xls.zip'
          excel = Excel.new(url, :zip)
          excel.default_sheet = excel.sheets.first
          assert_equal 'ist "e" im Nenner von H(s)', excel.cell('b', 5)
        ensure  
          excel.remove_tmp
        end  
      end
    end
  end

  def test_openoffice_open_from_uri_and_zipped
    if OPENOFFICE
      if ONLINE
        begin
          url = 'http://spazioinwind.libero.it/s2/rata.ods.zip'
          sheet = Openoffice.new(url, :zip)
          #has been changed: assert_equal 'ist "e" im Nenner von H(s)', sheet.cell('b', 5)
          assert_in_delta 0.001, 505.14, sheet.cell('c', 33).to_f
        ensure
          sheet.remove_tmp
        end  
      end
    end
  end

  def test_excel_zipped
    if EXCEL
      begin
        oo = Excel.new(File.join(TESTDIR,"bode-v1.xls.zip"), :zip)
        assert oo
        assert_equal 'ist "e" im Nenner von H(s)', oo.cell('b', 5)
      ensure
        oo.remove_tmp 
      end  
    end
  end

  def test_openoffice_zipped
    if OPENOFFICE
      begin
        oo = Openoffice.new(File.join(TESTDIR,"bode-v1.ods.zip"), :zip)
        assert oo
        assert_equal 'ist "e" im Nenner von H(s)', oo.cell('b', 5)
      ensure  
        oo.remove_tmp
      end  
    end
  end

  def test_bug_ric
    with_each_spreadsheet(:name=>'ric', :format=>:openoffice) do |oo|
      assert oo.empty?('A',1)
      assert oo.empty?('B',1)
      assert oo.empty?('C',1)
      assert oo.empty?('D',1)
      expected = 1
      letter = 'e'
      while letter <= 'u'
        assert_equal expected, oo.cell(letter,1)
        letter.succ!
        expected += 1
      end
      assert_equal 'J', oo.cell('v',1)
      assert_equal 'P', oo.cell('w',1)
      assert_equal 'B', oo.cell('x',1)
      assert_equal 'All', oo.cell('y',1)
      assert_equal 0, oo.cell('a',2)
      assert oo.empty?('b',2)
      assert oo.empty?('c',2)
      assert oo.empty?('d',2)
      assert_equal 'B', oo.cell('e',2)
      assert_equal 'B', oo.cell('f',2)
      assert_equal 'B', oo.cell('g',2)
      assert_equal 'B', oo.cell('h',2)
      assert_equal 'B', oo.cell('i',2)
      assert_equal 'B', oo.cell('j',2)
      assert_equal 'B', oo.cell('k',2)
      assert_equal 'B', oo.cell('l',2)
      assert_equal 'B', oo.cell('m',2)
      assert_equal 'B', oo.cell('n',2)
      assert_equal 'B', oo.cell('o',2)
      assert_equal 'B', oo.cell('p',2)
      assert_equal 'B', oo.cell('q',2)
      assert_equal 'B', oo.cell('r',2)
      assert_equal 'B', oo.cell('s',2)
      assert oo.empty?('t',2)
      assert oo.empty?('u',2)
      assert_equal 0  , oo.cell('v',2)
      assert_equal 0  , oo.cell('w',2)
      assert_equal 15 , oo.cell('x',2)
      assert_equal 15 , oo.cell('y',2)
    end
  end

  def test_mehrteilig
    with_each_spreadsheet(:name=>'Bibelbund1', :format=>:openoffice) do |oo|
      assert_equal "Tagebuch des Sekret\303\244rs.    Letzte Tagung 15./16.11.75 Schweiz", oo.cell(45,'A')
    end
    #if EXCELX
    #    #Datei gibt es noch nicht
    #    oo = Excelx.new(File.join(TESTDIR,"Bibelbund1.xlsx"))
    #    oo.default_sheet = oo.sheets.first
    #    assert_equal "Tagebuch des Sekret\303\244rs.    Letzte Tagung 15./16.11.75 Schweiz", oo.cell(45,'A')
    #end
  end

  def test_huge_document_to_csv
    if LONG_RUN
      with_each_spreadsheet(:name=>'Bibelbund') do |oo|
        assert_nothing_raised(Timeout::Error) {
          Timeout::timeout(GLOBAL_TIMEOUT) do |timeout_length|
            File.delete_if_exist("/tmp/Bibelbund.csv")
            assert_equal "Tagebuch des Sekret\303\244rs.    Letzte Tagung 15./16.11.75 Schweiz", oo.cell(45,'A')
            assert_equal "Tagebuch des Sekret\303\244rs.  Nachrichten aus Chile", oo.cell(46,'A')
            assert_equal "Tagebuch aus Chile  Juli 1977", oo.cell(55,'A')
            assert oo.to_csv("/tmp/Bibelbund.csv")
            assert File.exists?("/tmp/Bibelbund.csv")
            assert_equal "", `diff test/Bibelbund.csv /tmp/Bibelbund.csv`
          end 
        }
      end
    end
  end

  def test_to_csv
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      master = "#{TESTDIR}/numbers1.csv"
      File.delete_if_exist("/tmp/numbers1.csv")
      assert oo.to_csv("/tmp/numbers1.csv",oo.sheets.first)
      assert File.exists?("/tmp/numbers1.csv")
      assert_equal "", `diff #{master} /tmp/numbers1.csv`
      assert oo.to_csv("/tmp/numbers1.csv")
      assert File.exists?("/tmp/numbers1.csv")
      assert_equal "", `diff #{master} /tmp/numbers1.csv`
    end 
  end

  def test_bug_mehrere_datum
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      oo.default_sheet = 'Sheet5'
      assert_equal :date, oo.celltype('A',4)
      assert_equal :date, oo.celltype('B',4)
      assert_equal :date, oo.celltype('C',4)
      assert_equal :date, oo.celltype('D',4)
      assert_equal :date, oo.celltype('E',4)
      assert_equal Date.new(2007,11,21), oo.cell('A',4)
      assert_equal Date.new(2007,11,21), oo.cell('B',4)
      assert_equal Date.new(2007,11,21), oo.cell('C',4)
      assert_equal Date.new(2007,11,21), oo.cell('D',4)
      assert_equal Date.new(2007,11,21), oo.cell('E',4)
      assert_equal :float, oo.celltype('A',5)
      assert_equal :float, oo.celltype('B',5)
      assert_equal :float, oo.celltype('C',5)
      assert_equal :float, oo.celltype('D',5)
      assert_equal :float, oo.celltype('E',5)
      assert_equal 42, oo.cell('A',5)
      assert_equal 42, oo.cell('B',5)
      assert_equal 42, oo.cell('C',5)
      assert_equal 42, oo.cell('D',5)
      assert_equal 42, oo.cell('E',5)
      assert_equal :string, oo.celltype('A',6)
      assert_equal :string, oo.celltype('B',6)
      assert_equal :string, oo.celltype('C',6)
      assert_equal :string, oo.celltype('D',6)
      assert_equal :string, oo.celltype('E',6)
      assert_equal "ABC", oo.cell('A',6)
      assert_equal "ABC", oo.cell('B',6)
      assert_equal "ABC", oo.cell('C',6)
      assert_equal "ABC", oo.cell('D',6)
      assert_equal "ABC", oo.cell('E',6)
    end 
  end

  def test_multiple_sheets
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      2.times do
        oo.default_sheet = "Tabelle1"
        assert_equal 1, oo.cell(1,1)
        assert_equal 1, oo.cell(1,1,"Tabelle1")
        assert_equal "I am sheet 2", oo.cell('C',5,"Name of Sheet 2")
        sheetname = 'Sheet5'
        assert_equal :date, oo.celltype('A',4,sheetname)
        assert_equal :date, oo.celltype('B',4,sheetname)
        assert_equal :date, oo.celltype('C',4,sheetname)
        assert_equal :date, oo.celltype('D',4,sheetname)
        assert_equal :date, oo.celltype('E',4,sheetname)
        assert_equal Date.new(2007,11,21), oo.cell('A',4,sheetname)
        assert_equal Date.new(2007,11,21), oo.cell('B',4,sheetname)
        assert_equal Date.new(2007,11,21), oo.cell('C',4,sheetname)
        assert_equal Date.new(2007,11,21), oo.cell('D',4,sheetname)
        assert_equal Date.new(2007,11,21), oo.cell('E',4,sheetname)
        assert_equal :float, oo.celltype('A',5,sheetname)
        assert_equal :float, oo.celltype('B',5,sheetname)
        assert_equal :float, oo.celltype('C',5,sheetname)
        assert_equal :float, oo.celltype('D',5,sheetname)
        assert_equal :float, oo.celltype('E',5,sheetname)
        assert_equal 42, oo.cell('A',5,sheetname)
        assert_equal 42, oo.cell('B',5,sheetname)
        assert_equal 42, oo.cell('C',5,sheetname)
        assert_equal 42, oo.cell('D',5,sheetname)
        assert_equal 42, oo.cell('E',5,sheetname)
        assert_equal :string, oo.celltype('A',6,sheetname)
        assert_equal :string, oo.celltype('B',6,sheetname)
        assert_equal :string, oo.celltype('C',6,sheetname)
        assert_equal :string, oo.celltype('D',6,sheetname)
        assert_equal :string, oo.celltype('E',6,sheetname)
        assert_equal "ABC", oo.cell('A',6,sheetname)
        assert_equal "ABC", oo.cell('B',6,sheetname)
        assert_equal "ABC", oo.cell('C',6,sheetname)
        assert_equal "ABC", oo.cell('D',6,sheetname)
        assert_equal "ABC", oo.cell('E',6,sheetname)
        oo.reload
      end
    end
  end


  def test_bug_empty_sheet
    with_each_spreadsheet(:name=>'formula', :format=>[:openoffice, :excelx]) do |oo|
      oo.default_sheet = 'Sheet3' # is an empty sheet
      assert_nothing_raised(NoMethodError) {  oo.to_csv(File.join("/","tmp","emptysheet.csv"))  }
      assert_equal "", `cat /tmp/emptysheet.csv`
    end
  end

  def test_find_by_row_huge_document
    if LONG_RUN
      with_each_spreadsheet(:name=>'Bibelbund') do |oo|
        Timeout::timeout(GLOBAL_TIMEOUT) do |timeout_length|
          oo.default_sheet = oo.sheets.first
          rec = oo.find 20
          assert rec
          # assert_equal "Brief aus dem Sekretariat", rec[0]
          #p rec
          assert_equal "Brief aus dem Sekretariat", rec[0]['TITEL']
          rec = oo.find 22
          assert rec
          # assert_equal "Brief aus dem Skretariat. Tagung in Amberg/Opf.",rec[0]
          assert_equal "Brief aus dem Skretariat. Tagung in Amberg/Opf.",rec[0]['TITEL']
        end
      end
    end
  end

  def test_find_by_row
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      oo.header_line = nil
      rec = oo.find 16
      assert rec
      assert_nil oo.header_line
      # keine Headerlines in diesem Beispiel definiert
      assert_equal "einundvierzig", rec[0]
      #assert_equal false, rec
      rec = oo.find 15
      assert rec
      assert_equal 41,rec[0]
    end
  end

  def test_find_by_conditions
    if LONG_RUN
      with_each_spreadsheet(:name=>'Bibelbund') do |oo|
        assert_nothing_raised(Timeout::Error) {
          Timeout::timeout(GLOBAL_TIMEOUT) do |timeout_length|
            #-----------------------------------------------------------------
            zeilen = oo.find(:all, :conditions => {
                'TITEL' => 'Brief aus dem Sekretariat'
              }
            )
            assert_equal 2, zeilen.size
            assert_equal [{"VERFASSER"=>"Almassy, Annelene von",
                "INTERNET"=>nil,
                "SEITE"=>316.0,
                "KENNUNG"=>"Aus dem Bibelbund",
                "OBJEKT"=>"Bibel+Gem",
                "PC"=>"#C:\\Bibelbund\\reprint\\BuG1982-3.pdf#",
                "NUMMER"=>"1982-3",
                "TITEL"=>"Brief aus dem Sekretariat"},
              {"VERFASSER"=>"Almassy, Annelene von",
                "INTERNET"=>nil,
                "SEITE"=>222.0,
                "KENNUNG"=>"Aus dem Bibelbund",
                "OBJEKT"=>"Bibel+Gem",
                "PC"=>"#C:\\Bibelbund\\reprint\\BuG1983-2.pdf#",
                "NUMMER"=>"1983-2",
                "TITEL"=>"Brief aus dem Sekretariat"}] , zeilen

            #----------------------------------------------------------
            zeilen = oo.find(:all,
              :conditions => { 'VERFASSER' => 'Almassy, Annelene von' }
            )
            assert_equal 13, zeilen.size
            #----------------------------------------------------------
            zeilen = oo.find(:all, :conditions => {
                'TITEL' => 'Brief aus dem Sekretariat',
                'VERFASSER' => 'Almassy, Annelene von',
              }
            )
            assert_equal 2, zeilen.size
            assert_equal [{"VERFASSER"=>"Almassy, Annelene von",
                "INTERNET"=>nil,
                "SEITE"=>316.0,
                "KENNUNG"=>"Aus dem Bibelbund",
                "OBJEKT"=>"Bibel+Gem",
                "PC"=>"#C:\\Bibelbund\\reprint\\BuG1982-3.pdf#",
                "NUMMER"=>"1982-3",
                "TITEL"=>"Brief aus dem Sekretariat"},
              {"VERFASSER"=>"Almassy, Annelene von",
                "INTERNET"=>nil,
                "SEITE"=>222.0,
                "KENNUNG"=>"Aus dem Bibelbund",
                "OBJEKT"=>"Bibel+Gem",
                "PC"=>"#C:\\Bibelbund\\reprint\\BuG1983-2.pdf#",
                "NUMMER"=>"1983-2",
                "TITEL"=>"Brief aus dem Sekretariat"}] , zeilen

            # Result as an array
            zeilen = oo.find(:all,
              :conditions => {
                'TITEL' => 'Brief aus dem Sekretariat',
                'VERFASSER' => 'Almassy, Annelene von',
              }, :array => true)
            assert_equal 2, zeilen.size
            assert_equal [
              [
                "Brief aus dem Sekretariat",
                "Almassy, Annelene von",
                "Bibel+Gem",
                "1982-3",
                316.0,
                nil,
                "#C:\\Bibelbund\\reprint\\BuG1982-3.pdf#",
                "Aus dem Bibelbund",
              ],
              [
                "Brief aus dem Sekretariat",
                "Almassy, Annelene von",
                "Bibel+Gem",
                "1983-2",
                222.0,
                nil,
                "#C:\\Bibelbund\\reprint\\BuG1983-2.pdf#",
                "Aus dem Bibelbund",
              ]] , zeilen
          end # Timeout
        } # nothing_raised
      end
    end
  end


  #TODO: temporaerer Test
  def test_seiten_als_date
    with_each_spreadsheet(:name=>'Bibelbund', :format=>:excelx) do |oo|
      assert_equal 'Bericht aus dem Sekretariat', oo.cell(13,1)
      assert_equal '1981-4', oo.cell(13,'D')
      assert_equal [:numeric_or_formula,"General"], oo.excelx_type(13,'E')
      assert_equal '428', oo.excelx_value(13,'E')
      assert_equal 428.0, oo.cell(13,'E')
    end
  end

  def test_column
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      expected = [1.0,5.0,nil,10.0,Date.new(1961,11,21),'tata',nil,nil,nil,nil,'thisisa11',41.0,nil,nil,41.0,'einundvierzig',nil,Date.new(2007,5,31)]
      assert_equal expected, oo.column(1)
      assert_equal expected, oo.column('a')
    end
  end

  def test_column_huge_document
    if LONG_RUN
      with_each_spreadsheet(:name=>'Bibelbund') do |oo|
        assert_nothing_raised(Timeout::Error) {
          Timeout::timeout(GLOBAL_TIMEOUT) do |timeout_length|
            oo.default_sheet = oo.sheets.first
            assert_equal 3735, oo.column('a').size
            #assert_equal 499, oo.column('a').size
          end
        }
      end
    end
  end

  def test_simple_spreadsheet_find_by_condition
    with_each_spreadsheet(:name=>'simple_spreadsheet') do |oo|
      oo.header_line = 3
      oo.date_format = '%m/%d/%Y' if oo.class == Google
      erg = oo.find(:all, :conditions => {'Comment' => 'Task 1'})
      assert_equal Date.new(2007,05,07), erg[1]['Date']
      assert_equal 10.75       , erg[1]['Start time']
      assert_equal 12.50       , erg[1]['End time']
      assert_equal 0           , erg[1]['Pause']
      assert_equal 1.75        , erg[1]['Sum'] unless oo.class == Excel
      assert_equal "Task 1"    , erg[1]['Comment']
    end
  end
  
  # Ruby-spreadsheet now allows us to at least give the current value 
  # from a cell with a formula (no possible with parseexcel)
  def test_bug_false_borders_with_formulas
    with_each_spreadsheet(:name=>'false_encoding', :format=>:excel) do |oo|
      assert_equal 1, oo.first_row
      assert_equal 3, oo.last_row
      assert_equal 1, oo.first_column
      assert_equal 4, oo.last_column
    end
  end

  # We'ce added minimal formula support so we can now read these
  # though not sure how the spreadsheet reports older values....
  def test_fe
    with_each_spreadsheet(:name=>'false_encoding', :format=>:excel) do |oo|
      assert_equal Date.new(2007,11,1), oo.cell('a',1)
      #DOES NOT WORK IN EXCEL FILES: assert_equal true, oo.formula?('a',1)
      #DOES NOT WORK IN EXCEL FILES: assert_equal '=TODAY()', oo.formula('a',1)

      assert_equal Date.new(2008,2,9), oo.cell('B',1)
      #DOES NOT WORK IN EXCEL FILES: assert_equal true,               oo.formula?('B',1)
      #DOES NOT WORK IN EXCEL FILES: assert_equal "=A1+100",          oo.formula('B',1)

      assert_kind_of DateTime, oo.cell('C',1) 
      #DOES NOT WORK IN EXCEL FILES: assert_equal true,               oo.formula?('C',1)
      #DOES NOT WORK IN EXCEL FILES: assert_equal "=C1",          oo.formula('C',1)

      assert_equal 'H1', oo.cell('A',2)
      assert_equal 'H2', oo.cell('B',2)
      assert_equal 'H3', oo.cell('C',2)
      assert_equal 'H4', oo.cell('D',2)
      assert_equal 'R1', oo.cell('A',3)
      assert_equal 'R2', oo.cell('B',3)
      assert_equal 'R3', oo.cell('C',3)
      assert_equal 'R4', oo.cell('D',3)
    end  
  end

  def test_excel_does_not_support_formulas
    with_each_spreadsheet(:name=>'false_encoding', :format=>:excel) do |oo|
      assert_raise(RuntimeError) { void = oo.formula('a',1) }
      assert_raise(RuntimeError) { void = oo.formula?('a',1) }
      assert_raise(RuntimeError) { void = oo.formulas(oo.sheets.first) }
    end  
  end

  def get_extension(oo)
    case oo
    when Openoffice
      ".ods"
    when Excel
      ".xls"
    when Excelx
      ".xlsx"
    when Google  
      ""
    end
  end
  
  def test_info
    expected_templ = "File: numbers1%s\n"+
      "Number of sheets: 5\n"+
      "Sheets: Tabelle1, Name of Sheet 2, Sheet3, Sheet4, Sheet5\n"+
      "Sheet 1:\n"+
      "  First row: 1\n"+
      "  Last row: 18\n"+
      "  First column: A\n"+
      "  Last column: G\n"+
      "Sheet 2:\n"+
      "  First row: 5\n"+
      "  Last row: 14\n"+
      "  First column: B\n"+
      "  Last column: E\n"+
      "Sheet 3:\n"+
      "  First row: 1\n"+
      "  Last row: 1\n"+
      "  First column: A\n"+
      "  Last column: BA\n"+
      "Sheet 4:\n"+
      "  First row: 1\n"+
      "  Last row: 1\n"+
      "  First column: A\n"+
      "  Last column: E\n"+
      "Sheet 5:\n"+
      "  First row: 1\n"+
      "  Last row: 6\n"+
      "  First column: A\n"+
      "  Last column: E"
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      ext = get_extension(oo)
      expected = sprintf(expected_templ,ext)
      if oo.class == Google      
        assert_equal expected.gsub(/numbers1/,key_of("numbers1")), oo.info
      else
        assert_equal expected, oo.info
      end
    end
  end

  def test_bug_excel_numbers1_sheet5_last_row
    with_each_spreadsheet(:name=>'numbers1', :format=>:excel) do |oo|
      oo.default_sheet = "Tabelle1"
      assert_equal 1, oo.first_row
      assert_equal 18, oo.last_row
      assert_equal Openoffice.letter_to_number('A'), oo.first_column
      assert_equal Openoffice.letter_to_number('G'), oo.last_column
      oo.default_sheet = "Name of Sheet 2"
      assert_equal 5, oo.first_row
      assert_equal 14, oo.last_row
      assert_equal Openoffice.letter_to_number('B'), oo.first_column
      assert_equal Openoffice.letter_to_number('E'), oo.last_column
      oo.default_sheet = "Sheet3"
      assert_equal 1, oo.first_row
      assert_equal 1, oo.last_row
      assert_equal Openoffice.letter_to_number('A'), oo.first_column
      assert_equal Openoffice.letter_to_number('BA'), oo.last_column
      oo.default_sheet = "Sheet4"
      assert_equal 1, oo.first_row
      assert_equal 1, oo.last_row
      assert_equal Openoffice.letter_to_number('A'), oo.first_column
      assert_equal Openoffice.letter_to_number('E'), oo.last_column
      oo.default_sheet = "Sheet5"
      assert_equal 1, oo.first_row
      assert_equal 6, oo.last_row
      assert_equal Openoffice.letter_to_number('A'), oo.first_column
      assert_equal Openoffice.letter_to_number('E'), oo.last_column
    end
  end

  def test_should_raise_file_not_found_error
    if OPENOFFICE
      assert_raise(IOError) {
        oo = Openoffice.new(File.join('testnichtvorhanden','Bibelbund.ods'))
      }
    end
    if EXCEL
      assert_raise(IOError) {
        oo = Excel.new(File.join('testnichtvorhanden','Bibelbund.xls'))
      }
    end
    if EXCELX
      assert_raise(IOError) {
        oo = Excelx.new(File.join('testnichtvorhanden','Bibelbund.xlsx'))
      }
    end
    if GOOGLE
 #     assert_raise(Net::HTTPServerException) {
        # oo = Google.new(key_of('testnichtvorhanden'+'Bibelbund.ods'))
#        oo = Google.new('testnichtvorhanden')
#      }
    end
  end
  
  def test_write_google
    # write.me: http://spreadsheets.google.com/ccc?key=ptu6bbahNZpY0N0RrxQbWdw&hl=en_GB
    with_each_spreadsheet(:name=>'write.me', :format=>:google) do |oo|
      oo.set_value(1,1,"hello from the tests")
      assert_equal "hello from the tests", oo.cell(1,1)
      oo.set_value(1,1, 1.0)
      assert_equal 1.0, oo.cell(1,1)
    end
  end

  def test_bug_set_value_with_more_than_one_sheet_google
    # write.me: http://spreadsheets.google.com/ccc?key=ptu6bbahNZpY0N0RrxQbWdw&hl=en_GB
    with_each_spreadsheet(:name=>'write.me', :format=>:google) do |oo|
      content1 = 'AAA'
      content2 = 'BBB'
      oo.default_sheet = oo.sheets.first
      oo.set_value(1,1,content1)
      oo.default_sheet = oo.sheets[1]
      oo.set_value(1,1,content2) # in the second sheet
      oo.default_sheet = oo.sheets.first
      assert_equal content1, oo.cell(1,1)
      oo.default_sheet = oo.sheets[1]
      assert_equal content2, oo.cell(1,1)
    end
  end

  def test_set_value_with_sheet_argument_google
    with_each_spreadsheet(:name=>'write.me', :format=>:google) do |oo|
      random_row = rand(10)+1
      random_column = rand(10)+1
      content1 = 'ABC'
      content2 = 'DEF'
      oo.set_value(random_row,random_column,content1,oo.sheets.first)
      oo.set_value(random_row,random_column,content2,oo.sheets[1])
      assert_equal content1, oo.cell(random_row,random_column,oo.sheets.first)
      assert_equal content2, oo.cell(random_row,random_column,oo.sheets[1])
    end
  end

  def test_set_value_for_non_existing_sheet_google
    with_each_spreadsheet(:name=>'ptu6bbahNZpY0N0RrxQbWdw', :format=>:google) do |oo|
      assert_raise(RangeError) { oo.set_value(1,1,"dummy","no_sheet")   }
    end 
  end

  def test_bug_bbu
    with_each_spreadsheet(:name=>'bbu', :format=>[:openoffice, :excelx, :excel]) do |oo|
      assert_nothing_raised() {
      assert_equal "File: bbu#{get_extension(oo)}
Number of sheets: 3
Sheets: 2007_12, Tabelle2, Tabelle3
Sheet 1:
  First row: 1
  Last row: 4
  First column: A
  Last column: F
Sheet 2:
  - empty -
Sheet 3:
  - empty -", oo.info
    }

      oo.default_sheet = oo.sheets[1] # empty sheet
      assert_nil oo.first_row
      assert_nil oo.last_row
      assert_nil oo.first_column
      assert_nil oo.last_column
    end  
  end


  def test_bug_time_nil
    with_each_spreadsheet(:name=>'time-test') do |oo|
      assert_equal 12*3600+13*60+14, oo.cell('B',1) # 12:13:14 (secs since midnight)
      assert_equal :time, oo.celltype('B',1)
      assert_equal 15*3600+16*60, oo.cell('C',1) # 15:16    (secs since midnight)
      assert_equal :time, oo.celltype('C',1)
      assert_equal 23*3600, oo.cell('D',1) # 23:00    (secs since midnight)
      assert_equal :time, oo.celltype('D',1)
    end
  end

  def test_date_time_to_csv
    with_each_spreadsheet(:name=>'time-test') do |oo|
      begin
        assert oo.to_csv("/tmp/time-test.csv")
        assert File.exists?("/tmp/time-test.csv")
        assert_equal "", `diff #{TESTDIR}/time-test.csv /tmp/time-test.csv`
      ensure
        File.delete_if_exist("/tmp/time-test.csv")
      end  
    end 
  end

  def test_date_time_yaml
    with_each_spreadsheet(:name=>'time-test') do |oo|
      expected =
        "--- \ncell_1_1: \n  row: 1 \n  col: 1 \n  celltype: string \n  value: Mittags: \ncell_1_2: \n  row: 1 \n  col: 2 \n  celltype: time \n  value: 12:13:14 \ncell_1_3: \n  row: 1 \n  col: 3 \n  celltype: time \n  value: 15:16:00 \ncell_1_4: \n  row: 1 \n  col: 4 \n  celltype: time \n  value: 23:00:00 \ncell_2_1: \n  row: 2 \n  col: 1 \n  celltype: date \n  value: 2007-11-21 \n"
      assert_equal expected, oo.to_yaml
    end
  end

  def test_no_remaining_tmp_files_openoffice
    if OPENOFFICE
      assert_raise(Zip::ZipError) { #TODO: besseres Fehlerkriterium bei
        # oo = Openoffice.new(File.join(TESTDIR,"no_spreadsheet_file.txt"))
        # es soll absichtlich ein Abbruch provoziert werden, deshalb :ignore
        oo = Openoffice.new(File.join(TESTDIR,"no_spreadsheet_file.txt"),
          false,
          :ignore)
      }
      a=Dir.glob("oo_*")
      assert_equal [], a
    end
  end

  def test_no_remaining_tmp_files_excel
    if EXCEL
      assert_raise(Ole::Storage::FormatError) {
        # oo = Excel.new(File.join(TESTDIR,"no_spreadsheet_file.txt"))
        # es soll absichtlich ein Abbruch provoziert werden, deshalb :ignore
        oo = Excel.new(File.join(TESTDIR,"no_spreadsheet_file.txt"),
          false,
          :ignore)
      }
      a=Dir.glob("oo_*")
      assert_equal [], a
    end
  end

  def test_no_remaining_tmp_files_excelx
    if EXCELX
      assert_raise(Zip::ZipError) { #TODO: besseres Fehlerkriterium bei

        # oo = Excelx.new(File.join(TESTDIR,"no_spreadsheet_file.txt"))
        # es soll absichtlich ein Abbruch provoziert werden, deshalb :ignore
        oo = Excelx.new(File.join(TESTDIR,"no_spreadsheet_file.txt"),
          false,
          :ignore)

      }
      a=Dir.glob("oo_*")
      assert_equal [], a
    end
  end

  def test_no_remaining_tmp_files_google
    if GOOGLE
      assert_nothing_raised() {
        oo = Google.new(key_of("no_spreadsheet_file.txt"))
      }
      a=Dir.glob("oo_*")
      assert_equal [], a
    end
  end

  # Erstellt eine Liste aller Zellen im Spreadsheet. Dies ist nötig, weil ein einfacher
  # Textvergleich des XML-Outputs nicht funktioniert, da xml-builder die Attribute
  # nicht immer in der gleichen Reihenfolge erzeugt.
  def init_all_cells(oo,sheet)
    all = []
    oo.first_row(sheet).upto(oo.last_row(sheet)) do |row|
      oo.first_column(sheet).upto(oo.last_column(sheet)) do |col|
        unless oo.empty?(row,col,sheet)
          all << {:row => row.to_s,
            :column => col.to_s,
            :content => oo.cell(row,col,sheet).to_s,
            :type => oo.celltype(row,col,sheet).to_s,
          }
        end
      end
    end
    all
  end

  def test_to_xml
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      assert_nothing_raised {oo.to_xml}
      sheetname = oo.sheets.first
      doc = XML::Parser.string(oo.to_xml).parse
      doc.root.each_element {|xml_sheet|
        all_cells = init_all_cells(oo, sheetname)
        x = 0
        assert_equal sheetname, xml_sheet.attributes['name']
        xml_sheet.each_element {|cell|
          expected = [all_cells[x][:row],
            all_cells[x][:column],
            all_cells[x][:content],
            all_cells[x][:type],
          ]
          result = [
            cell.attributes['row'],
            cell.attributes['column'],
            cell.content,
            cell.attributes['type'],
          ]
          assert_equal expected, result
          x += 1
        } # end of sheet
        sheetname = oo.sheets[oo.sheets.index(sheetname)+1]
      }
    end  
  end

  def test_bug_row_column_fixnum_float
    with_each_spreadsheet(:name=>'bug-row-column-fixnum-float', :format=>:excel) do |oo|
      assert_equal 42.5, oo.cell('b',2)
      assert_equal 43  , oo.cell('c',2)
      assert_equal ['hij',42.5, 43], oo.row(2)
      assert_equal ['def',42.5, 'nop'], oo.column(2)
    end
  end

  def test_file_warning_default
    if OPENOFFICE
      assert_raises(TypeError) { oo = Openoffice.new(File.join(TESTDIR,"numbers1.xls")) }
      assert_raises(TypeError) { oo = Openoffice.new(File.join(TESTDIR,"numbers1.xlsx")) }
      assert_equal [], Dir.glob("oo_*")
    end
    if EXCEL
      assert_raises(TypeError) { oo = Excel.new(File.join(TESTDIR,"numbers1.ods")) }
      assert_raises(TypeError) { oo = Excel.new(File.join(TESTDIR,"numbers1.xlsx")) }
      assert_equal [], Dir.glob("oo_*")
    end
    if EXCELX
      assert_raises(TypeError) { oo = Excelx.new(File.join(TESTDIR,"numbers1.ods")) }
      assert_raises(TypeError) { oo = Excelx.new(File.join(TESTDIR,"numbers1.xls")) }
      assert_equal [], Dir.glob("oo_*")
    end
  end

  def test_file_warning_error
    if OPENOFFICE
      assert_raises(TypeError) { oo = Openoffice.new(File.join(TESTDIR,"numbers1.xls"),false,:error) }
      assert_raises(TypeError) { oo = Openoffice.new(File.join(TESTDIR,"numbers1.xlsx"),false,:error) }
      assert_equal [], Dir.glob("oo_*")
    end
    if EXCEL
      assert_raises(TypeError) { oo = Excel.new(File.join(TESTDIR,"numbers1.ods"),false,:error) }
      assert_raises(TypeError) { oo = Excel.new(File.join(TESTDIR,"numbers1.xlsx"),false,:error) }
      assert_equal [], Dir.glob("oo_*")
    end
    if EXCELX
      assert_raises(TypeError) { oo = Excelx.new(File.join(TESTDIR,"numbers1.ods"),false,:error) }
      assert_raises(TypeError) { oo = Excelx.new(File.join(TESTDIR,"numbers1.xls"),false,:error) }
      assert_equal [], Dir.glob("oo_*")
    end
  end

  def test_file_warning_warning
    if OPENOFFICE
      assert_nothing_raised(TypeError) {
        assert_raises(Zip::ZipError) {
          oo = Openoffice.new(File.join(TESTDIR,"numbers1.xls"),false, :warning)
        }
      }
      assert_nothing_raised(TypeError) {
        assert_raises(Errno::ENOENT) {
          oo = Openoffice.new(File.join(TESTDIR,"numbers1.xlsx"),false, :warning)
        }
      }
      assert_equal [], Dir.glob("oo_*")
    end
    if EXCEL
      assert_nothing_raised(TypeError) {
        assert_raises(Ole::Storage::FormatError) {
          oo = Excel.new(File.join(TESTDIR,"numbers1.ods"),false, :warning) }
      }
      assert_nothing_raised(TypeError) {
        assert_raises(Ole::Storage::FormatError) {
          oo = Excel.new(File.join(TESTDIR,"numbers1.xlsx"),false, :warning) }
      }
      assert_equal [], Dir.glob("oo_*")
    end
    if EXCELX
      assert_nothing_raised(TypeError) {
        assert_raises(Errno::ENOENT) {
          oo = Excelx.new(File.join(TESTDIR,"numbers1.ods"),false, :warning) }
      }
      assert_nothing_raised(TypeError) {
        assert_raises(Zip::ZipError) {
          oo = Excelx.new(File.join(TESTDIR,"numbers1.xls"),false, :warning) }
      }
      assert_equal [], Dir.glob("oo_*")
    end
  end

  def test_file_warning_ignore
    if OPENOFFICE
      assert_nothing_raised(TypeError) {
        assert_raises(Zip::ZipError) {
          oo = Openoffice.new(File.join(TESTDIR,"numbers1.xls"),false, :ignore) }
      }
      assert_nothing_raised(TypeError) {
        assert_raises(Errno::ENOENT) {
          oo = Openoffice.new(File.join(TESTDIR,"numbers1.xlsx"),false, :ignore) }
      }
      assert_equal [], Dir.glob("oo_*")
    end
    if EXCEL
      assert_nothing_raised(TypeError) {
        assert_raises(Ole::Storage::FormatError) {
          oo = Excel.new(File.join(TESTDIR,"numbers1.ods"),false, :ignore) }
      }
      assert_nothing_raised(TypeError) {
        assert_raises(Ole::Storage::FormatError) {oo = Excel.new(File.join(TESTDIR,"numbers1.xlsx"),false, :ignore) }}
      assert_equal [], Dir.glob("oo_*")
    end
    if EXCELX
      assert_nothing_raised(TypeError) {
        assert_raises(Errno::ENOENT) {
          oo = Excelx.new(File.join(TESTDIR,"numbers1.ods"),false, :ignore)
        }
      }
      assert_nothing_raised(TypeError) {
        assert_raises(Zip::ZipError) {
          oo = Excelx.new(File.join(TESTDIR,"numbers1.xls"),false, :ignore)
        }
      }
      assert_equal [], Dir.glob("oo_*")
    end
  end

  def test_bug_last_row_excel
    with_each_spreadsheet(:name=>'time-test', :format=>:excel) do |oo|    
      assert_equal 2, oo.last_row
    end
  end

  def test_bug_to_xml_with_empty_sheets
    with_each_spreadsheet(:name=>'emptysheets', :format=>[:openoffice, :excel]) do |oo|    
      oo.sheets.each { |sheet|
        assert_equal nil, oo.first_row, "first_row not nil in sheet #{sheet}"
        assert_equal nil, oo.last_row, "last_row not nil in sheet #{sheet}"
        assert_equal nil, oo.first_column, "first_column not nil in sheet #{sheet}"
        assert_equal nil, oo.last_column, "last_column not nil in sheet #{sheet}"
        assert_equal nil, oo.first_row(sheet), "first_row not nil in sheet #{sheet}"
        assert_equal nil, oo.last_row(sheet), "last_row not nil in sheet #{sheet}"
        assert_equal nil, oo.first_column(sheet), "first_column not nil in sheet #{sheet}"
        assert_equal nil, oo.last_column(sheet), "last_column not nil in sheet #{sheet}"
      }
      assert_nothing_raised() { result = oo.to_xml }
    end
  end

  def test_bug_simple_spreadsheet_time_bug
    # really a bug? are cells really of type time?
    # No! :float must be the correct type
    with_each_spreadsheet(:name=>'simple_spreadsheet', :format=>:excelx) do |oo|    
      # puts oo.cell('B',5).to_s
      # assert_equal :time, oo.celltype('B',5)
      assert_equal :float, oo.celltype('B',5)
      assert_equal 10.75, oo.cell('B',5)
      assert_equal 12.50, oo.cell('C',5)
      assert_equal 0, oo.cell('D',5)
      assert_equal 1.75, oo.cell('E',5)
      assert_equal 'Task 1', oo.cell('F',5)
      assert_equal Date.new(2007,5,7), oo.cell('A',5)
    end
  end

  def test_simple2_excelx
    with_each_spreadsheet(:name=>'simple_spreadsheet', :format=>:excelx) do |oo|    
      assert_equal [:numeric_or_formula, "yyyy\\-mm\\-dd"], oo.excelx_type('A',4)
      assert_equal [:numeric_or_formula, "#,##0.00"], oo.excelx_type('B',4)
      assert_equal [:numeric_or_formula, "#,##0.00"], oo.excelx_type('c',4)
      assert_equal [:numeric_or_formula, "General"], oo.excelx_type('d',4)
      assert_equal [:numeric_or_formula, "General"], oo.excelx_type('e',4)
      assert_equal :string, oo.excelx_type('f',4)

      assert_equal "39209", oo.excelx_value('a',4)
      assert_equal "yyyy\\-mm\\-dd", oo.excelx_format('a',4)
      assert_equal "9.25", oo.excelx_value('b',4)
      assert_equal "10.25", oo.excelx_value('c',4)
      assert_equal "0", oo.excelx_value('d',4)
      #... Sum-Spalte
      # assert_equal "Task 1", oo.excelx_value('f',4)
      assert_equal "Task 1", oo.cell('f',4)
      assert_equal Date.new(2007,05,07), oo.cell('a',4)
      assert_equal "9.25", oo.excelx_value('b',4)
      assert_equal "#,##0.00", oo.excelx_format('b',4)
      assert_equal 9.25, oo.cell('b',4)
      assert_equal :float, oo.celltype('b',4)
      assert_equal :float, oo.celltype('d',4)
      assert_equal 0, oo.cell('d',4)
      assert_equal :formula, oo.celltype('e',4)
      assert_equal 1, oo.cell('e',4)
      assert_equal 'C4-B4-D4', oo.formula('e',4)
      assert_equal :string, oo.celltype('f',4)
      assert_equal "Task 1", oo.cell('f',4)
    end
  end

  def test_datetime
    with_each_spreadsheet(:name=>'datetime') do |oo|    
      val = oo.cell('c',3)
      assert_kind_of DateTime, val
      assert_equal :datetime, oo.celltype('c',3)
      assert_equal DateTime.new(1961,11,21,12,17,18), val
      val = oo.cell('a',1)
      assert_kind_of Date, val
      assert_equal :date, oo.celltype('a',1)
      assert_equal Date.new(1961,11,21), val
      assert_equal Date.new(1961,11,21), oo.cell('a',1)
      assert_equal DateTime.new(1961,11,21,12,17,18), oo.cell('a',3)
      assert_equal DateTime.new(1961,11,21,12,17,18), oo.cell('b',3)
      assert_equal DateTime.new(1961,11,21,12,17,18), oo.cell('c',3)
      assert_equal DateTime.new(1961,11,21,12,17,18), oo.cell('a',4)
      assert_equal DateTime.new(1961,11,21,12,17,18), oo.cell('b',4)
      assert_equal DateTime.new(1961,11,21,12,17,18), oo.cell('c',4)
      assert_equal DateTime.new(1961,11,21,12,17,18), oo.cell('a',5)
      assert_equal DateTime.new(1961,11,21,12,17,18), oo.cell('b',5)
      assert_equal DateTime.new(1961,11,21,12,17,18), oo.cell('c',5)
      assert_equal Date.new(1961,11,21), oo.cell('a',6)
      assert_equal Date.new(1961,11,21), oo.cell('b',6)
      assert_equal Date.new(1961,11,21), oo.cell('c',6)
      assert_equal Date.new(1961,11,21), oo.cell('a',7)
      assert_equal Date.new(1961,11,21), oo.cell('b',7)
      assert_equal Date.new(1961,11,21), oo.cell('c',7)
    end
  end
  
  def test_cell_openoffice_html_escape
    with_each_spreadsheet(:name=>'html-escape', :format=>:openoffice) do |oo|    
      assert_equal "'", oo.cell(1,1)
      assert_equal "&", oo.cell(2,1)
      assert_equal ">", oo.cell(3,1)
      assert_equal "<", oo.cell(4,1)
      assert_equal "`", oo.cell(5,1)
      # test_openoffice_zipped will catch issues with &quot;
    end  
  end    
  
  def test_cell_boolean
    with_each_spreadsheet(:name=>'boolean', :format=>[:openoffice, :excel, :excelx]) do |oo|    
      if oo.class == Excelx    
        assert_equal "TRUE", oo.cell(1,1)
        assert_equal "FALSE", oo.cell(2,1)
      else
        assert_equal "true", oo.cell(1,1)
        assert_equal "false", oo.cell(2,1)
      end
    end
  end
  
   def test_cell_multiline
     with_each_spreadsheet(:name=>'paragraph', :format=>[:openoffice, :excel, :excelx]) do |oo|    
        assert_equal "This is a test\nof a multiline\nCell", oo.cell(1,1)
        assert_equal "This is a test\n¶\nof a multiline\n\nCell", oo.cell(1,2)
        assert_equal "first p\n\nsecond p\n\nlast p", oo.cell(2,1)
     end
   end
  
  def test_cell_styles
    with_each_spreadsheet(:name=>'style', :format=>[:openoffice, :excel, :excelx]) do |oo|    
      # bold
      assert_equal true,  oo.font(1,1).bold?
      assert_equal false, oo.font(1,1).italic?
      assert_equal false, oo.font(1,1).underline?
    
      # italic
      assert_equal false, oo.font(2,1).bold?
      assert_equal true,  oo.font(2,1).italic?
      assert_equal false, oo.font(2,1).underline?

      # normal
      assert_equal false, oo.font(3,1).bold?
      assert_equal false, oo.font(3,1).italic?
      assert_equal false, oo.font(3,1).underline?

      # underline
      assert_equal false, oo.font(4,1).bold?
      assert_equal false, oo.font(4,1).italic?
      assert_equal true,  oo.font(4,1).underline?

      # bold italic
      assert_equal true,  oo.font(5,1).bold?
      assert_equal true,  oo.font(5,1).italic?
      assert_equal false, oo.font(5,1).underline?

      # bold underline
      assert_equal true,  oo.font(6,1).bold?
      assert_equal false, oo.font(6,1).italic?
      assert_equal true,  oo.font(6,1).underline?

      # italic underline
      assert_equal false, oo.font(7,1).bold?
      assert_equal true,  oo.font(7,1).italic?
      assert_equal true,  oo.font(7,1).underline?

      # bolded row
      assert_equal true, oo.font(8,1).bold?
      assert_equal false,  oo.font(8,1).italic?
      assert_equal false,  oo.font(8,1).underline?

      # bolded col
      assert_equal true, oo.font(9,2).bold?
      assert_equal false,  oo.font(9,2).italic?
      assert_equal false,  oo.font(9,2).underline?

      # bolded row, italic col
      assert_equal true, oo.font(10,3).bold?
      assert_equal true,  oo.font(10,3).italic?
      assert_equal false,  oo.font(10,3).underline?

      # normal
      assert_equal false, oo.font(11,4).bold?
      assert_equal false,  oo.font(11,4).italic?
      assert_equal false,  oo.font(11,4).underline?
    end  
  end
  
  # If a cell has a date-like string but is preceeded by a ' 
  # to force that date to be treated like a string, we were getting an exception.
  # This test just checks for that exception to make sure it's not raised in this case
  def test_date_to_float_conversion
    with_each_spreadsheet(:name=>'datetime_floatconv', :format=>:excel) do |oo|    
       assert_nothing_raised(NoMethodError) do
         oo.cell('a',1)
         oo.cell('a',2)
       end
     end
  end  

# Need to extend to other formats
  def test_row_whitespace
    with_each_spreadsheet(:name=>'whitespace') do |oo| 
      oo.default_sheet = "Sheet1"
      assert_equal [nil, nil, nil, nil, nil, nil], oo.row(1)
      assert_equal [nil, nil, nil, nil, nil, nil], oo.row(2)
      assert_equal ["Date", "Start time", "End time", "Pause", "Sum", "Comment"], oo.row(3)
      assert_equal [Date.new(2007,5,7), 9.25, 10.25, 0.0, 1.0, "Task 1"], oo.row(4)
      assert_equal [nil, nil, nil, nil, nil, nil], oo.row(5)
      assert_equal [Date.new(2007,5,7), 10.75, 10.75, 0.0, 0.0, "Task 1"], oo.row(6)
      oo.default_sheet = "Sheet2"
      assert_equal ["Date", nil, "Start time"], oo.row(1)
      assert_equal [Date.new(2007,5,7), nil, 9.25], oo.row(2)
      assert_equal [Date.new(2007,5,7), nil,  10.75], oo.row(3)
    end
  end
  
  def test_col_whitespace
    with_each_spreadsheet(:name=>'whitespace') do |oo|    
      oo.default_sheet = "Sheet1"
      assert_equal ["Date", Date.new(2007,5,7), nil, Date.new(2007,5,7)], oo.column(1)
      assert_equal ["Start time", 9.25, nil, 10.75], oo.column(2)
      assert_equal ["End time", 10.25, nil, 10.75], oo.column(3)
      assert_equal ["Pause", 0.0, nil, 0.0], oo.column(4)
      assert_equal ["Sum", 1.0, nil, 0.0], oo.column(5)
      assert_equal ["Comment","Task 1", nil, "Task 1"], oo.column(6)
      oo.default_sheet = "Sheet2"
      assert_equal [nil, nil, nil], oo.column(1)
      assert_equal [nil, nil, nil], oo.column(2)
      assert_equal ["Date", Date.new(2007,5,7), Date.new(2007,5,7)], oo.column(3)
      assert_equal [nil, nil, nil], oo.column(4)
      assert_equal [ "Start time", 9.25, 10.75], oo.column(5)
    end
  end
  

  def test_ruby_spreadsheet_formula_bug
     with_each_spreadsheet(:name=>'formula_parse_error', :format=>:excel) do |oo|
       assert_equal '5026', oo.cell(2,3)
       assert_equal '5026', oo.cell(3,3)
     end
   end


  # Excel has two base date formats one from 1900 and the other from 1904. 
  # There's a MS bug that 1900 base dates include an extra day due to erroneously
  # including 1900 as a leap yar. 
  def test_base_dates_in_excel
    with_each_spreadsheet(:name=>'1900_base', :format=>:excel) do |oo|    
      assert_equal Date.new(2009,06,15), oo.cell(1,1)
      assert_equal Date.new(2009,06,28), oo.cell(2,1) #formula for TODAY(), last calculated on 06.28
      assert_equal :date, oo.celltype(1,1)
    end  
    with_each_spreadsheet(:name=>'1904_base', :format=>:excel) do |oo|    
      assert_equal Date.new(2009,06,15), oo.cell(1,1)
      assert_equal Date.new(2009,06,28), oo.cell(2,1) #formula for TODAY(), last calculated on 06.28
      assert_equal :date, oo.celltype(1,1)
    end  
  end
   
end # class
