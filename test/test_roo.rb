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
require File.dirname(__FILE__) + '/test_helper.rb'
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
        'numbers1' => "o10837434939102457526.4784396906364855777",
        'borders' => "o10837434939102457526.664868920231926255",
        'simple_spreadsheet' => "ptu6bbahNZpYe-L1vEBmgGA",
        'testnichtvorhandenBibelbund.ods' => "invalidkeyforanyspreadsheet", # !!! intentionally false key
        "only_one_sheet" => "o10837434939102457526.762705759906130135",
        "write.me" => 'ptu6bbahNZpY0N0RrxQbWdw&hl',
        'formula' => 'o10837434939102457526.3022866619437760118',
        'time-test' => 'ptu6bbahNZpYBMhk01UfXSg',
        'datetime' => "ptu6bbahNZpYQEtZwzL_dZQ",
      }[spreadsheetname]
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

  OPENOFFICE = true  	# do Openoffice-Spreadsheet Tests?
  EXCEL      = true	# do Excel Tests?
  GOOGLE     = false 	# do Google-Spreadsheet Tests?
  GNUMERIC_ODS = false # do gnumeric with ods files Tests?
  EXCELX      = true  	# do Excel-X Tests? (.xlsx-files)

  OPENOFFICEWRITE = false # experimental: write access with OO-Documents
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

  def test_date
    assert Google.date?("21/11/1962")
    assert_equal Date.new(1962,11,21), Google.to_date("21/11/1962")

    assert !Google.date?("21")
    assert_nil Google.to_date("21")

    assert !Google.date?("21/11")
    assert_nil Google.to_date("21/11")

    assert !Google.date?("Mittwoch/21/1961")
    assert_nil Google.to_date("Mittwoch/21/1961")
  end

  def test_classes
    if OPENOFFICE
      oo = Openoffice.new(File.join("test","numbers1.ods"))
      assert_kind_of Openoffice, oo
    end
    if EXCEL
      oo = Excel.new(File.join("test","numbers1.xls"))
      assert_kind_of Excel, oo
    end
    if GOOGLE
      oo = Google.new(key_of("numbers1"))
      assert_kind_of Google, oo
    end
    if EXCELX
      oo = Excelx.new(File.join("test","numbers1.xlsx"))
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

  def DONT_test_simple_google
    if GOOGLE
      go = Google.new("egal")
      assert_equal "42", go.cell(1,1)
    end
  end

  def test_sheets_openoffice
    if OPENOFFICE
      oo = Openoffice.new(File.join("test","numbers1.ods"))
      assert_equal ["Tabelle1","Name of Sheet 2","Sheet3","Sheet4","Sheet5"], oo.sheets
      assert_raise(RangeError) {
        oo.default_sheet = "no_sheet"
      }
      assert_raise(TypeError) {
        oo.default_sheet = [1,2,3]
      }

      oo.sheets.each { |sh|
        oo.default_sheet = sh
        assert_equal sh, oo.default_sheet
      }
    end
  end

  def test_sheets_gnumeric_ods
    if GNUMERIC_ODS
      oo = Openoffice.new(File.join("test","gnumeric_numbers1.ods"))
      assert_equal ["Tabelle1","Name of Sheet 2","Sheet3","Sheet4","Sheet5"], oo.sheets
      assert_raise(RangeError) {
        oo.default_sheet = "no_sheet"
      }
      assert_raise(TypeError) {
        oo.default_sheet = [1,2,3]
      }

      oo.sheets.each { |sh|
        oo.default_sheet = sh
        assert_equal sh, oo.default_sheet
      }
    end
  end

  def test_sheets_excel
    if EXCEL
      oo = Excel.new(File.join("test","numbers1.xls"))
      assert_equal ["Tabelle1","Name of Sheet 2","Sheet3","Sheet4","Sheet5"], oo.sheets
      assert_raise(RangeError) {
        oo.default_sheet = "no_sheet"
      }
      assert_raise(TypeError) {
        oo.default_sheet = [1,2,3]
      }
      oo.sheets.each { |sh|
        oo.default_sheet = sh
        assert_equal sh, oo.default_sheet
      }
    end
  end

  def test_sheets_excelx
    if EXCELX
      oo = Excelx.new(File.join("test","numbers1.xlsx"))
      assert_equal ["Tabelle1","Name of Sheet 2","Sheet3","Sheet4","Sheet5"], oo.sheets
      assert_raise(RangeError) {
        oo.default_sheet = "no_sheet"
      }
      assert_raise(TypeError) {
        oo.default_sheet = [1,2,3]
      }
      oo.sheets.each { |sh|
        oo.default_sheet = sh
        assert_equal sh, oo.default_sheet
      }
    end
  end

  def test_sheets_google
    if GOOGLE
      oo = Google.new(key_of("numbers1"))
      assert_equal ["Tabelle1","Name of Sheet 2","Sheet3","Sheet4","Sheet5"], oo.sheets
      assert_raise(RangeError) {
        oo.default_sheet = "no_sheet"
      }
      assert_raise(TypeError) {
        oo.default_sheet = [1,2,3]
      }
      oo.sheets.each { |sh|
        oo.default_sheet = sh
        assert_equal sh, oo.default_sheet
      }
    end
  end

  def test_cell_openoffice
    if OPENOFFICE
      oo = Openoffice.new(File.join("test","numbers1.ods"))
      oo.default_sheet = oo.sheets.first
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
      # assert_equal "string", oo.celltype(2,6)
      assert_equal :string, oo.celltype(2,6)
      assert_equal 11, oo.cell(2,7)
      # assert_equal "float", oo.celltype(2,7)
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

      # assert_equal "date", oo.celltype(5,1)
      assert_equal :date, oo.celltype(5,1)
      assert_equal Date.new(1961,11,21), oo.cell(5,1)
      assert_equal "1961-11-21", oo.cell(5,1).to_s
    end
  end

  def test_cell_gnumeric_ods
    if GNUMERIC_ODS
      oo = Openoffice.new(File.join("test","gnumeric_numbers1.ods"))
      oo.default_sheet = oo.sheets.first
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
      # assert_equal "string", oo.celltype(2,6)
      assert_equal :string, oo.celltype(2,6)
      assert_equal 11, oo.cell(2,7)
      # assert_equal "float", oo.celltype(2,7)
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

      # assert_equal "date", oo.celltype(5,1)
      assert_equal :date, oo.celltype(5,1)
      assert_equal Date.new(1961,11,21), oo.cell(5,1)
      assert_equal "1961-11-21", oo.cell(5,1).to_s
    end
  end

  def test_cell_excel
    if EXCEL
      oo = Excel.new(File.join("test","numbers1.xls"))
      oo.default_sheet = oo.sheets.first
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
      # assert_equal "string", oo.celltype(2,6)
      assert_equal :string, oo.celltype(2,6)
      assert_equal 11, oo.cell(2,7)
      # assert_equal "float", oo.celltype(2,7)
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

      # assert_equal "date", oo.celltype(5,1)
      assert_equal :date, oo.celltype(5,1)
      assert_equal Date.new(1961,11,21), oo.cell(5,1)
      assert_equal "1961-11-21", oo.cell(5,1).to_s
    end
  end

  def test_cell_excelx
    if EXCELX
      oo = Excelx.new(File.join("test","numbers1.xlsx"))
      oo.default_sheet = oo.sheets.first

      assert_kind_of Float, oo.cell(1,1)
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
      # assert_equal "string", oo.celltype(2,6)
      assert_equal :string, oo.celltype(2,6)
      assert_equal 11, oo.cell(2,7)
      # assert_equal "float", oo.celltype(2,7)
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

      # assert_equal "date", oo.celltype(5,1)
      assert_equal :date, oo.celltype(5,1)
      assert_equal Date.new(1961,11,21), oo.cell(5,1)
      assert_equal "1961-11-21", oo.cell(5,1).to_s
    end
  end

  def test_cell_google
    if GOOGLE
      oo = Google.new(key_of("numbers1"))
      oo.default_sheet = oo.sheets.first
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
      # assert_equal "string", oo.celltype(2,6)
      assert_equal :string, oo.celltype(2,6)
      assert_equal 11, oo.cell(2,7)
      # assert_equal "float", oo.celltype(2,7)
      assert_equal :float, oo.celltype(2,7), "Inhalt: --#{oo.cell(2,7)}--"

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

      # assert_equal "date", oo.celltype(5,1)
      assert_equal :date, oo.celltype(5,1)
      assert_equal Date.new(1961,11,21), oo.cell(5,1)
      assert_equal "1961-11-21", oo.cell(5,1).to_s
    end # GOOGLE
  end

  def test_celltype
    ###
    if OPENOFFICE
      oo = Openoffice.new(File.join("test","numbers1.ods"))
      oo.default_sheet = oo.sheets.first
      assert_equal :string, oo.celltype(2,6)
    end
    if GNUMERIC_ODS
      oo = Openoffice.new(File.join("test","gnumeric_numbers1.ods"))
      oo.default_sheet = oo.sheets.first
      assert_equal :string, oo.celltype(2,6)
    end
    if EXCEL
      oo = Excel.new(File.join("test","numbers1.xls"))
      oo.default_sheet = oo.sheets.first
      assert_equal :string, oo.celltype(2,6)
    end
    if GOOGLE
      oo = Google.new(key_of("numbers1"))
      oo.default_sheet = oo.sheets.first
      assert_equal :string, oo.celltype(2,6)
    end
    if EXCELX
      oo = Excelx.new(File.join("test","numbers1.xlsx"))
      oo.default_sheet = oo.sheets.first
      assert_equal :string, oo.celltype(2,6)
    end
  end

  def test_cell_address
    if OPENOFFICE
      oo = Openoffice.new(File.join("test","numbers1.ods"))
      oo.default_sheet = oo.sheets.first
      assert_equal "tata", oo.cell(6,1)
      assert_equal "tata", oo.cell(6,'A')
      assert_equal "tata", oo.cell('A',6)
      assert_equal "tata", oo.cell(6,'a')
      assert_equal "tata", oo.cell('a',6)

      assert_raise(ArgumentError) {
        assert_equal "tata", oo.cell('a','f')
      }
      assert_raise(ArgumentError) {
        assert_equal "tata", oo.cell('f','a')
      }
      assert_equal "thisisc8", oo.cell(8,3)
      assert_equal "thisisc8", oo.cell(8,'C')
      assert_equal "thisisc8", oo.cell('C',8)
      assert_equal "thisisc8", oo.cell(8,'c')
      assert_equal "thisisc8", oo.cell('c',8)

      assert_equal "thisisd9", oo.cell('d',9)
      assert_equal "thisisa11", oo.cell('a',11)
    end

    if GNUMERIC_ODS
      oo = Openoffice.new(File.join("test","gnumeric_numbers1.ods"))
      oo.default_sheet = oo.sheets.first
      assert_equal "tata", oo.cell(6,1)
      assert_equal "tata", oo.cell(6,'A')
      assert_equal "tata", oo.cell('A',6)
      assert_equal "tata", oo.cell(6,'a')
      assert_equal "tata", oo.cell('a',6)

      assert_raise(ArgumentError) {
        assert_equal "tata", oo.cell('a','f')
      }
      assert_raise(ArgumentError) {
        assert_equal "tata", oo.cell('f','a')
      }
      assert_equal "thisisc8", oo.cell(8,3)
      assert_equal "thisisc8", oo.cell(8,'C')
      assert_equal "thisisc8", oo.cell('C',8)
      assert_equal "thisisc8", oo.cell(8,'c')
      assert_equal "thisisc8", oo.cell('c',8)

      assert_equal "thisisd9", oo.cell('d',9)
      assert_equal "thisisa11", oo.cell('a',11)
    end

    if EXCEL
      oo = Excel.new(File.join("test","numbers1.xls"))
      oo.default_sheet = oo.sheets.first
      assert_equal "tata", oo.cell(6,'A')
      assert_equal "tata", oo.cell(6,1)
      assert_equal "tata", oo.cell('A',6)
      assert_equal "tata", oo.cell(6,'a')
      assert_equal "tata", oo.cell('a',6)

      assert_raise(ArgumentError) {
        assert_equal "tata", oo.cell('a','f')
      }
      assert_raise(ArgumentError) {
        assert_equal "tata", oo.cell('f','a')
      }

      assert_equal "thisisc8", oo.cell(8,3)
      assert_equal "thisisc8", oo.cell(8,'C')
      assert_equal "thisisc8", oo.cell('C',8)
      assert_equal "thisisc8", oo.cell(8,'c')
      assert_equal "thisisc8", oo.cell('c',8)

      assert_equal "thisisd9", oo.cell('d',9)
      assert_equal "thisisa11", oo.cell('a',11)
      #assert_equal "lulua", oo.cell('b',10)
    end

    if EXCELX
      oo = Excelx.new(File.join("test","numbers1.xlsx"))
      oo.default_sheet = oo.sheets.first
      assert_equal "tata", oo.cell(6,1)
      assert_equal "tata", oo.cell(6,'A')
      assert_equal "tata", oo.cell('A',6)
      assert_equal "tata", oo.cell(6,'a')
      assert_equal "tata", oo.cell('a',6)

      assert_raise(ArgumentError) {
        assert_equal "tata", oo.cell('a','f')
      }
      assert_raise(ArgumentError) {
        assert_equal "tata", oo.cell('f','a')
      }

      assert_equal "thisisc8", oo.cell(8,3)
      assert_equal "thisisc8", oo.cell(8,'C')
      assert_equal "thisisc8", oo.cell('C',8)
      assert_equal "thisisc8", oo.cell(8,'c')
      assert_equal "thisisc8", oo.cell('c',8)

      assert_equal "thisisd9", oo.cell('d',9)
      assert_equal "thisisa11", oo.cell('a',11)
      #assert_equal "lulua", oo.cell('b',10)
    end

    if GOOGLE
      oo = Google.new(key_of("numbers1"))
      oo.default_sheet = oo.sheets.first
      assert_equal "tata", oo.cell(6,1)
      assert_equal "tata", oo.cell(6,'A')
      assert_equal "tata", oo.cell('A',6)
      assert_equal "tata", oo.cell(6,'a')
      assert_equal "tata", oo.cell('a',6)

      assert_raise(ArgumentError) {
        assert_equal "tata", oo.cell('a','f')
      }
      assert_raise(ArgumentError) {
        assert_equal "tata", oo.cell('f','a')
      }
      assert_equal "thisisc8", oo.cell(8,3)
      assert_equal "thisisc8", oo.cell(8,'C')
      assert_equal "thisisc8", oo.cell('C',8)
      assert_equal "thisisc8", oo.cell(8,'c')
      assert_equal "thisisc8", oo.cell('c',8)

      assert_equal "thisisd9", oo.cell('d',9)
      assert_equal "thisisa11", oo.cell('a',11)
    end
  end

  # Version of the (XML) office document
  # please note that "1.0" is returned even if it was created with OpenOffice V. 2.0
  def test_officeversion
    if OPENOFFICE
      oo = Openoffice.new(File.join("test","numbers1.ods"))
      assert_equal "1.0", oo.officeversion
    end
    if GNUMERIC_ODS
      oo = Openoffice.new(File.join("test","gnumeric_numbers1.ods"))
      assert_equal "1.0", oo.officeversion
    end
    if EXCEL
      # excel does not have a officeversion
    end
    if EXCELX
      # excelx does not have a officeversion
      #TODO: gibt es hier eine Versionsnummer
    end
    if GOOGLE
      # google does not have a officeversion
    end
  end

  #TODO: inkonsequente Lieferung Fixnum/Float
  def test_rows
    if OPENOFFICE
      oo = Openoffice.new(File.join("test","numbers1.ods"))
      oo.default_sheet = oo.sheets.first
      assert_equal 41, oo.cell('a',12)
      assert_equal 42, oo.cell('b',12)
      assert_equal 43, oo.cell('c',12)
      assert_equal 44, oo.cell('d',12)
      assert_equal 45, oo.cell('e',12)
      assert_equal [41.0,42.0,43.0,44.0,45.0], oo.row(12)
      assert_equal "einundvierzig", oo.cell('a',16)
      assert_equal "zweiundvierzig", oo.cell('b',16)
      assert_equal "dreiundvierzig", oo.cell('c',16)
      assert_equal "vierundvierzig", oo.cell('d',16)
      assert_equal "fuenfundvierzig", oo.cell('e',16)
      assert_equal ["einundvierzig", "zweiundvierzig", "dreiundvierzig", "vierundvierzig", "fuenfundvierzig"], oo.row(16)
    end
    if GNUMERIC_ODS
      oo = Openoffice.new(File.join("test","gnumeric_numbers1.ods"))
      oo.default_sheet = oo.sheets.first
      assert_equal 41, oo.cell('a',12)
      assert_equal 42, oo.cell('b',12)
      assert_equal 43, oo.cell('c',12)
      assert_equal 44, oo.cell('d',12)
      assert_equal 45, oo.cell('e',12)
      assert_equal [41.0,42.0,43.0,44.0,45.0], oo.row(12)
      assert_equal "einundvierzig", oo.cell('a',16)
      assert_equal "zweiundvierzig", oo.cell('b',16)
      assert_equal "dreiundvierzig", oo.cell('c',16)
      assert_equal "vierundvierzig", oo.cell('d',16)
      assert_equal "fuenfundvierzig", oo.cell('e',16)
      assert_equal ["einundvierzig", "zweiundvierzig", "dreiundvierzig", "vierundvierzig", "fuenfundvierzig"], oo.row(16)
    end
    if EXCEL
      oo = Excel.new(File.join("test","numbers1.xls"))
      oo.default_sheet = oo.sheets.first
      assert_equal 41, oo.cell('a',12)
      assert_equal 42, oo.cell('b',12)
      assert_equal 43, oo.cell('c',12)
      assert_equal 44, oo.cell('d',12)
      assert_equal 45, oo.cell('e',12)
      assert_equal [41,42,43,44,45], oo.row(12)
      assert_equal "einundvierzig", oo.cell('a',16)
      assert_equal "zweiundvierzig", oo.cell('b',16)
      assert_equal "dreiundvierzig", oo.cell('c',16)
      assert_equal "vierundvierzig", oo.cell('d',16)
      assert_equal "fuenfundvierzig", oo.cell('e',16)
      assert_equal ["einundvierzig",
        "zweiundvierzig",
        "dreiundvierzig",
        "vierundvierzig",
        "fuenfundvierzig"], oo.row(16)
    end
    if EXCELX
      oo = Excelx.new(File.join("test","numbers1.xlsx"))
      oo.default_sheet = oo.sheets.first
      assert_equal 41, oo.cell('a',12)
      assert_equal 42, oo.cell('b',12)
      assert_equal 43, oo.cell('c',12)
      assert_equal 44, oo.cell('d',12)
      assert_equal 45, oo.cell('e',12)
      assert_equal [41,42,43,44,45], oo.row(12)
      assert_equal "einundvierzig", oo.cell('a',16)
      assert_equal "zweiundvierzig", oo.cell('b',16)
      assert_equal "dreiundvierzig", oo.cell('c',16)
      assert_equal "vierundvierzig", oo.cell('d',16)
      assert_equal "fuenfundvierzig", oo.cell('e',16)
      assert_equal ["einundvierzig",
        "zweiundvierzig",
        "dreiundvierzig",
        "vierundvierzig",
        "fuenfundvierzig"], oo.row(16)
    end
    if GOOGLE
      oo = Google.new(key_of("numbers1"))
      oo.default_sheet = oo.sheets.first
      assert_equal 41, oo.cell('a',12)
      assert_equal 42, oo.cell('b',12)
      assert_equal 43, oo.cell('c',12)
      assert_equal 44, oo.cell('d',12)
      assert_equal 45, oo.cell('e',12)
      assert_equal [41,42,43,44,45], oo.row(12)
      assert_equal "einundvierzig", oo.cell('a',16)
      assert_equal "zweiundvierzig", oo.cell('b',16)
      assert_equal "dreiundvierzig", oo.cell('c',16)
      assert_equal "vierundvierzig", oo.cell('d',16)
      assert_equal "fuenfundvierzig", oo.cell('e',16)
      assert_equal ["einundvierzig",
        "zweiundvierzig",
        "dreiundvierzig",
        "vierundvierzig",
        "fuenfundvierzig"], oo.row(16)
    end
  end

  def test_last_row
    if OPENOFFICE
      oo = Openoffice.new(File.join("test","numbers1.ods"))
      oo.default_sheet = oo.sheets.first
      assert_equal 18, oo.last_row
    end
    if GNUMERIC_ODS
      oo = Openoffice.new(File.join("test","gnumeric_numbers1.ods"))
      oo.default_sheet = oo.sheets.first
      assert_equal 18, oo.last_row
    end
    if EXCEL
      oo = Excel.new(File.join("test","numbers1.xls"))
      oo.default_sheet = oo.sheets.first
      assert_equal 18, oo.last_row
    end
    if EXCELX
      oo = Excelx.new(File.join("test","numbers1.xlsx"))
      oo.default_sheet = oo.sheets.first
      assert_equal 18, oo.last_row
    end
    if GOOGLE
      oo = Google.new(key_of("numbers1"))
      oo.default_sheet = oo.sheets.first
      assert_equal 18, oo.last_row
    end
  end

  def test_last_column
    if OPENOFFICE
      oo = Openoffice.new(File.join("test","numbers1.ods"))
      oo.default_sheet = oo.sheets.first
      assert_equal 7, oo.last_column
    end
    if GNUMERIC_ODS
      oo = Openoffice.new(File.join("test","gnumeric_numbers1.ods"))
      oo.default_sheet = oo.sheets.first
      assert_equal 7, oo.last_column
    end
    if EXCEL
      oo = Excel.new(File.join("test","numbers1.xls"))
      oo.default_sheet = oo.sheets.first
      assert_equal 7, oo.last_column
    end
    if EXCELX
      oo = Excelx.new(File.join("test","numbers1.xlsx"))
      oo.default_sheet = oo.sheets.first
      assert_equal 7, oo.last_column
    end
    if GOOGLE
      oo = Google.new(key_of("numbers1"))
      oo.default_sheet = oo.sheets.first
      assert_equal 7, oo.last_column
    end
  end

  def test_last_column_as_letter
    if OPENOFFICE
      oo = Openoffice.new(File.join("test","numbers1.ods"))
      oo.default_sheet = oo.sheets.first
      assert_equal 'G', oo.last_column_as_letter
    end
    if GNUMERIC_ODS
      oo = Openoffice.new(File.join("test","gnumeric_numbers1.ods"))
      oo.default_sheet = oo.sheets.first
      assert_equal 'G', oo.last_column_as_letter
    end
    if EXCEL
      oo = Excel.new(File.join("test","numbers1.xls"))
      oo.default_sheet = 1 # oo.sheets.first
      assert_equal 'G', oo.last_column_as_letter
    end
    if EXCELX
      oo = Excelx.new(File.join("test","numbers1.xlsx"))
      oo.default_sheet = oo.sheets.first
      assert_equal 'G', oo.last_column_as_letter
    end
    if GOOGLE
      oo = Google.new(key_of("numbers1"))
      oo.default_sheet = oo.sheets.first
      assert_equal 'G', oo.last_column_as_letter
    end
  end

  def test_first_row
    if OPENOFFICE
      oo = Openoffice.new(File.join("test","numbers1.ods"))
      oo.default_sheet = oo.sheets.first
      assert_equal 1, oo.first_row
    end
    if GNUMERIC_ODS
      oo = Openoffice.new(File.join("test","gnumeric_numbers1.ods"))
      oo.default_sheet = oo.sheets.first
      assert_equal 1, oo.first_row
    end
    if EXCEL
      oo = Excel.new(File.join("test","numbers1.xls"))
      oo.default_sheet = 1 # oo.sheets.first
      assert_equal 1, oo.first_row
    end
    if EXCELX
      oo = Excelx.new(File.join("test","numbers1.xlsx"))
      oo.default_sheet = oo.sheets.first
      assert_equal 1, oo.first_row
    end
    if GOOGLE
      oo = Google.new(key_of("numbers1"))
      oo.default_sheet = oo.sheets.first
      assert_equal 1, oo.first_row
    end
  end

  def test_first_column
    if OPENOFFICE
      oo = Openoffice.new(File.join("test","numbers1.ods"))
      oo.default_sheet = oo.sheets.first
      assert_equal 1, oo.first_column
    end
    if GNUMERIC_ODS
      oo = Openoffice.new(File.join("test","gnumeric_numbers1.ods"))
      oo.default_sheet = oo.sheets.first
      assert_equal 1, oo.first_column
    end
    if EXCEL
      oo = Excel.new(File.join("test","numbers1.xls"))
      oo.default_sheet = 1 # oo.sheets.first
      assert_equal 1, oo.first_column
    end
    if EXCELX
      oo = Excelx.new(File.join("test","numbers1.xlsx"))
      oo.default_sheet = oo.sheets.first
      assert_equal 1, oo.first_column
    end
    if GOOGLE
      assert_nothing_raised(Timeout::Error) {
        Timeout::timeout(GLOBAL_TIMEOUT) do |timeout_length|
          oo = Google.new(key_of("numbers1"))
          oo.default_sheet = oo.sheets.first
          assert_equal 1, oo.first_column
        end
      }
    end
  end

  def test_first_column_as_letter_openoffice
    if OPENOFFICE
      oo = Openoffice.new(File.join("test","numbers1.ods"))
      oo.default_sheet = oo.sheets.first
      assert_equal 'A', oo.first_column_as_letter
    end
  end

  def test_first_column_as_letter_gnumeric_ods
    if GNUMERIC_ODS
      oo = Openoffice.new(File.join("test","gnumeric_numbers1.ods"))
      oo.default_sheet = oo.sheets.first
      assert_equal 'A', oo.first_column_as_letter
    end
  end

  def test_first_column_as_letter_excel
    if EXCEL
      oo = Excel.new(File.join("test","numbers1.xls"))
      oo.default_sheet = 1 # oo.sheets.first
      assert_equal 'A', oo.first_column_as_letter
    end
  end

  def test_first_column_as_letter_excelx
    if EXCELX
      oo = Excelx.new(File.join("test","numbers1.xlsx"))
      oo.default_sheet = oo.sheets.first
      assert_equal 'A', oo.first_column_as_letter
    end
  end

  def test_first_column_as_letter_google
    if GOOGLE
      oo = Google.new(key_of("numbers1"))
      oo.default_sheet = oo.sheets.first
      assert_equal 'A', oo.first_column_as_letter
    end
  end

  def test_sheetname
    if OPENOFFICE
      oo = Openoffice.new(File.join("test","numbers1.ods"))
      oo.default_sheet = "Name of Sheet 2"
      assert_equal 'I am sheet 2', oo.cell('C',5)
      assert_raise(RangeError) { oo.default_sheet = "non existing sheet name" }
      assert_raise(RangeError) { dummy = oo.cell('C',5,"non existing sheet name")}
      assert_raise(RangeError) { dummy = oo.celltype('C',5,"non existing sheet name")}
      assert_raise(RangeError) { dummy = oo.empty?('C',5,"non existing sheet name")}
      assert_raise(RangeError) { dummy = oo.formula?('C',5,"non existing sheet name")}
      assert_raise(RangeError) { dummy = oo.formula('C',5,"non existing sheet name")}
      assert_raise(RangeError) { dummy = oo.set('C',5,42,"non existing sheet name")}
      assert_raise(RangeError) { dummy = oo.formulas("non existing sheet name")}
      assert_raise(RangeError) { dummy = oo.to_yaml({},1,1,1,1,"non existing sheet name")}
    end
    if GNUMERIC_ODS
      oo = Openoffice.new(File.join("test","gnumeric_numbers1.ods"))
      oo.default_sheet = "Name of Sheet 2"
      assert_equal 'I am sheet 2', oo.cell('C',5)
      assert_raise(RangeError) {
        oo.default_sheet = "non existing sheet name"
      }
      assert_raise(RangeError) { oo.default_sheet = "non existing sheet name" }
      assert_raise(RangeError) { dummy = oo.cell('C',5,"non existing sheet name")}
      assert_raise(RangeError) { dummy = oo.celltype('C',5,"non existing sheet name")}
      assert_raise(RangeError) { dummy = oo.empty?('C',5,"non existing sheet name")}
      assert_raise(RangeError) { dummy = oo.formula?('C',5,"non existing sheet name")}
      assert_raise(RangeError) { dummy = oo.formula('C',5,"non existing sheet name")}
      assert_raise(RangeError) { dummy = oo.set('C',5,42,"non existing sheet name")}
      assert_raise(RangeError) { dummy = oo.formulas("non existing sheet name")}
      assert_raise(RangeError) { dummy = oo.to_yaml({},1,1,1,1,"non existing sheet name")}
    end
    if EXCEL
      oo = Excel.new(File.join("test","numbers1.xls"))
      oo.default_sheet = "Name of Sheet 2"
      assert_equal 'I am sheet 2', oo.cell('C',5)
      assert_raise(RangeError) {
        oo.default_sheet = "non existing sheet name"
      }
      assert_raise(RangeError) { oo.default_sheet = "non existing sheet name" }
      assert_raise(RangeError) { dummy = oo.cell('C',5,"non existing sheet name")}
      assert_raise(RangeError) { dummy = oo.celltype('C',5,"non existing sheet name")}
      assert_raise(RangeError) { dummy = oo.empty?('C',5,"non existing sheet name")}
      assert_raise(RuntimeError) { dummy = oo.formula?('C',5,"non existing sheet name")}
      assert_raise(RuntimeError) { dummy = oo.formula('C',5,"non existing sheet name")}
      #assert_raise(RangeError) { dummy = oo.set('C',5,42,"non existing sheet name")}
      #assert_raise(RangeError) { dummy = oo.formulas("non existing sheet name")}
      assert_raise(RangeError) { dummy = oo.to_yaml({},1,1,1,1,"non existing sheet name")}
    end
    if EXCELX
      oo = Excelx.new(File.join("test","numbers1.xlsx"))
      oo.default_sheet = "Name of Sheet 2"
      assert_equal 'I am sheet 2', oo.cell('C',5)
      assert_raise(RangeError) {
        oo.default_sheet = "non existing sheet name"
      }
      assert_raise(RangeError) { oo.default_sheet = "non existing sheet name" }
      assert_raise(RangeError) { dummy = oo.cell('C',5,"non existing sheet name")}
      assert_raise(RangeError) { dummy = oo.celltype('C',5,"non existing sheet name")}
      assert_raise(RangeError) { dummy = oo.empty?('C',5,"non existing sheet name")}
      assert_raise(RangeError) { dummy = oo.formula?('C',5,"non existing sheet name")}
      assert_raise(RangeError) { dummy = oo.formula('C',5,"non existing sheet name")}
      assert_raise(RangeError) { dummy = oo.set('C',5,42,"non existing sheet name")}
      assert_raise(RangeError) { dummy = oo.formulas("non existing sheet name")}
      assert_raise(RangeError) { dummy = oo.to_yaml({},1,1,1,1,"non existing sheet name")}
    end
    if GOOGLE
      oo = Google.new(key_of("numbers1"))
      oo.default_sheet = "Name of Sheet 2"
      assert_equal 'I am sheet 2', oo.cell('C',5)
      assert_raise(RangeError) {
        oo.default_sheet = "non existing sheet name"
      }
      assert_raise(RangeError) { oo.default_sheet = "non existing sheet name" }
      assert_raise(RangeError) { dummy = oo.cell('C',5,"non existing sheet name")}
      assert_raise(RangeError) { dummy = oo.celltype('C',5,"non existing sheet name")}
      assert_raise(RangeError) { dummy = oo.empty?('C',5,"non existing sheet name")}
      assert_raise(RangeError) { dummy = oo.formula?('C',5,"non existing sheet name")}
      assert_raise(RangeError) { dummy = oo.formula('C',5,"non existing sheet name")}
      #2008-12-04: assert_raise(RangeError) { dummy = oo.set('C',5,42,"non existing sheet name")}
      assert_raise(RangeError) { dummy = oo.set_value('C',5,42,"non existing sheet name")}
      assert_raise(RangeError) { dummy = oo.formulas("non existing sheet name")}
      assert_raise(RangeError) { dummy = oo.to_yaml({},1,1,1,1,"non existing sheet name")}
    end
  end

  def test_boundaries
    if OPENOFFICE
      oo = Openoffice.new(File.join("test","numbers1.ods"))
      oo.default_sheet = "Name of Sheet 2"
      assert_equal 2, oo.first_column
      assert_equal 'B', oo.first_column_as_letter
      assert_equal 5, oo.first_row
      assert_equal 'E', oo.last_column_as_letter
      assert_equal 14, oo.last_row
    end
    if GNUMERIC_ODS
      oo = Openoffice.new(File.join("test","gnumeric_numbers1.ods"))
      oo.default_sheet = "Name of Sheet 2"
      assert_equal 2, oo.first_column
      assert_equal 'B', oo.first_column_as_letter
      assert_equal 5, oo.first_row
      assert_equal 'E', oo.last_column_as_letter
      assert_equal 14, oo.last_row
    end
    if EXCEL
      #-- Excel
      oo = Excel.new(File.join("test","numbers1.xls"))
      oo.default_sheet = 2 # "Name of Sheet 2"
      assert_equal 2, oo.first_column
      assert_equal 'B', oo.first_column_as_letter
      assert_equal 5, oo.first_row
      assert_equal 'E', oo.last_column_as_letter
      assert_equal 14, oo.last_row
    end
    if EXCELX
      oo = Excelx.new(File.join("test","numbers1.xlsx"))
      oo.default_sheet = "Name of Sheet 2"
      assert_equal 2, oo.first_column
      assert_equal 'B', oo.first_column_as_letter
      assert_equal 5, oo.first_row
      assert_equal 'E', oo.last_column_as_letter
      assert_equal 14, oo.last_row
    end
  end

  def test_multiple_letters
    if OPENOFFICE
      oo = Openoffice.new(File.join("test","numbers1.ods"))
      oo.default_sheet = "Sheet3"
      assert_equal "i am AA", oo.cell('AA',1)
      assert_equal "i am AB", oo.cell('AB',1)
      assert_equal "i am BA", oo.cell('BA',1)
      assert_equal 'BA', oo.last_column_as_letter
      assert_equal "i am BA", oo.cell(1,'BA')
    end
    if GNUMERIC_ODS
      oo = Openoffice.new(File.join("test","gnumeric_numbers1.ods"))
      oo.default_sheet = "Sheet3"
      assert_equal "i am AA", oo.cell('AA',1)
      assert_equal "i am AB", oo.cell('AB',1)
      assert_equal "i am BA", oo.cell('BA',1)
      assert_equal 'BA', oo.last_column_as_letter
      assert_equal "i am BA", oo.cell(1,'BA')
    end
    if EXCEL
      oo = Excel.new(File.join("test","numbers1.xls"))
      oo.default_sheet = 3 # "Sheet3"
      assert_equal "i am AA", oo.cell('AA',1)
      assert_equal "i am AB", oo.cell('AB',1)
      assert_equal "i am BA", oo.cell('BA',1)
      assert_equal 'BA', oo.last_column_as_letter
      assert_equal "i am BA", oo.cell(1,'BA')
    end
    if EXCELX
      oo = Excelx.new(File.join("test","numbers1.xlsx"))
      oo.default_sheet = "Sheet3"
      assert_equal "i am AA", oo.cell('AA',1)
      assert_equal "i am AB", oo.cell('AB',1)
      assert_equal "i am BA", oo.cell('BA',1)
      assert_equal 'BA', oo.last_column_as_letter
      assert_equal "i am BA", oo.cell(1,'BA')
    end
  end

  def test_argument_error
    if EXCEL
      oo = Excel.new(File.join("test","numbers1.xls"))
      before Date.new(2007,7,20) do
        assert_raise(ArgumentError) {
          #oo.default_sheet = "first sheet"
          oo.default_sheet = "Tabelle1"
        }
      end
      assert_nothing_raised(ArgumentError) {
        # oo.default_sheet = 1
        #oo.default_sheet = "first sheet"
        oo.default_sheet = "Tabelle1"
      }
    end
  end

  def test_empty_eh
    if OPENOFFICE #-- OpenOffice
      oo = Openoffice.new(File.join("test","numbers1.ods"))
      oo.default_sheet = oo.sheets.first
      assert oo.empty?('a',14)
      assert ! oo.empty?('a',15)
      assert oo.empty?('a',20)
    end
    if GNUMERIC_ODS
      oo = Openoffice.new(File.join("test","gnumeric_numbers1.ods"))
      oo.default_sheet = oo.sheets.first
      assert oo.empty?('a',14)
      assert ! oo.empty?('a',15)
      assert oo.empty?('a',20)
    end
    if EXCEL
      oo = Excel.new(File.join("test","numbers1.xls"))
      oo.default_sheet = 1
      assert oo.empty?('a',14)
      assert ! oo.empty?('a',15)
      assert oo.empty?('a',20)
    end
    if EXCELX
      oo = Excelx.new(File.join("test","numbers1.xlsx"))
      oo.default_sheet = oo.sheets.first
      assert oo.empty?('a',14)
      assert ! oo.empty?('a',15)
      assert oo.empty?('a',20)
    end
  end

  def test_writeopenoffice
    if OPENOFFICEWRITE
      File.cp(File.join("test","numbers1.ods"),
        File.join("test","numbers2.ods"))
      File.cp(File.join("test","numbers2.ods"),
        File.join("test","bak_numbers2.ods"))
      oo = Openoffice.new(File.join("test","numbers2.ods"))
      oo.default_sheet = oo.sheets.first
      oo.first_row.upto(oo.last_row) {|y|
        oo.first_column.upto(oo.last_column) {|x|
          unless oo.empty?(y,x)
            # oo.set(y, x, oo.cell(y,x) + 7) if oo.celltype(y,x) == "float"
            oo.set(y, x, oo.cell(y,x) + 7) if oo.celltype(y,x) == :float
          end
        }
      }
      oo.save

      oo1 = Openoffice.new(File.join("test","numbers2.ods"))
      oo2 = Openoffice.new(File.join("test","bak_numbers2.ods"))
      #p oo2.to_s
      assert_equal 999, oo2.cell('a',1), oo2.cell('a',1)
      assert_equal oo2.cell('a',1) + 7, oo1.cell('a',1)
      assert_equal oo2.cell('b',1)+7, oo1.cell('b',1)
      assert_equal oo2.cell('c',1)+7, oo1.cell('c',1)
      assert_equal oo2.cell('d',1)+7, oo1.cell('d',1)
      assert_equal oo2.cell('a',2)+7, oo1.cell('a',2)
      assert_equal oo2.cell('b',2)+7, oo1.cell('b',2)
      assert_equal oo2.cell('c',2)+7, oo1.cell('c',2)
      assert_equal oo2.cell('d',2)+7, oo1.cell('d',2)
      assert_equal oo2.cell('e',2)+7, oo1.cell('e',2)

      File.cp(File.join("test","bak_numbers2.ods"),
        File.join("test","numbers2.ods"))
    end
  end

  def test_reload
    if OPENOFFICE
      oo = Openoffice.new(File.join("test","numbers1.ods"))
      oo.default_sheet = oo.sheets.first
      assert_equal 1, oo.cell(1,1)

      oo.reload
      assert_equal 1, oo.cell(1,1)
    end
    if GNUMERIC_ODS
      oo = Openoffice.new(File.join("test","gnumeric_numbers1.ods"))
      oo.default_sheet = oo.sheets.first
      assert_equal 1, oo.cell(1,1)

      oo.reload
      assert_equal 1, oo.cell(1,1)
    end
    if EXCEL
      oo = Excel.new(File.join("test","numbers1.xls"))
      oo.default_sheet = 1 # oo.sheets.first
      assert_equal 1, oo.cell(1,1)

      oo.reload
      assert_equal 1, oo.cell(1,1)
    end
    if EXCELX
      oo = Excelx.new(File.join("test","numbers1.xlsx"))
      oo.default_sheet = oo.sheets.first
      assert_equal 1, oo.cell(1,1)

      oo.reload
      assert_equal 1, oo.cell(1,1)
    end
  end

  def test_bug_contiguous_cells
    if OPENOFFICE
      oo = Openoffice.new(File.join("test","numbers1.ods"))
      oo.default_sheet = "Sheet4"
      assert_equal Date.new(2007,06,16), oo.cell('a',1)
      assert_equal 10, oo.cell('b',1)
      assert_equal 10, oo.cell('c',1)
      assert_equal 10, oo.cell('d',1)
      assert_equal 10, oo.cell('e',1)
    end
    if GNUMERIC_ODS
      oo = Openoffice.new(File.join("test","gnumeric_numbers1.ods"))
      oo.default_sheet = "Sheet4"
      assert_equal Date.new(2007,06,16), oo.cell('a',1)
      assert_equal 10, oo.cell('b',1)
      assert_equal 10, oo.cell('c',1)
      assert_equal 10, oo.cell('d',1)
      assert_equal 10, oo.cell('e',1)
    end
    if EXCEL
      # dieser Test ist fuer Excel sheets nicht noetig,
      # da der Bug nur bei OO-Dokumenten auftrat
    end
    if GOOGLE
      # dieser Test ist fuer Google sheets nicht noetig,
      # da der Bug nur bei OO-Dokumenten auftrat
    end
  end

  def test_bug_italo_ve
    if OPENOFFICE
      oo = Openoffice.new(File.join("test","numbers1.ods"))
      oo.default_sheet = "Sheet5"
      assert_equal 1, oo.cell('A',1)
      assert_equal 5, oo.cell('b',1)
      assert_equal 5, oo.cell('c',1)
      assert_equal 2, oo.cell('a',2)
      assert_equal 3, oo.cell('a',3)
    end
    if OPENOFFICE
      oo = Openoffice.new(File.join("test","numbers1.ods"))
      oo.default_sheet = "Sheet5"
      assert_equal 1, oo.cell('A',1)
      assert_equal 5, oo.cell('b',1)
      assert_equal 5, oo.cell('c',1)
      assert_equal 2, oo.cell('a',2)
      assert_equal 3, oo.cell('a',3)
    end
    if GNUMERIC_ODS
      oo = Openoffice.new(File.join("test","gnumeric_numbers1.ods"))
      oo.default_sheet = "Sheet5"
      assert_equal 1, oo.cell('A',1)
      assert_equal 5, oo.cell('b',1)
      assert_equal 5, oo.cell('c',1)
      assert_equal 2, oo.cell('a',2)
      assert_equal 3, oo.cell('a',3)
    end
    if EXCEL
      oo = Excel.new(File.join("test","numbers1.xls"))
      oo.default_sheet = 5
      assert_equal 1, oo.cell('A',1)
      assert_equal 5, oo.cell('b',1)
      assert_equal 5, oo.cell('c',1)
      assert_equal 2, oo.cell('a',2)
      assert_equal 3, oo.cell('a',3)
    end
    if EXCELX
      oo = Excelx.new(File.join("test","numbers1.xlsx"))
      oo.default_sheet = "Sheet5" # oo.sheets[5-1]
      assert_equal 1, oo.cell('A',1)
      assert_equal 5, oo.cell('b',1)
      assert_equal 5, oo.cell('c',1)
      assert_equal 2, oo.cell('a',2)
      assert_equal 3, oo.cell('a',3)
    end
    #if GOOGLE
    #  oo = Google.new(key_of("numbers1"))
    #  oo.default_sheet = "Sheet5"
    #  assert_equal 1, oo.cell('A',1)
    #  assert_equal 5, oo.cell('b',1)
    #  assert_equal 5, oo.cell('c',1)
    #  assert_equal 2, oo.cell('a',2)
    #  assert_equal 3, oo.cell('a',3)
    #end
  end

  #2008-01-30
  def test_italo_table
    local_only do
      oo = Openoffice.new(File.join("test","simple_spreadsheet_from_italo.ods"))
      oo.default_sheet = oo.sheets.first

      assert_equal  '1', oo.cell('A',1)
      assert_equal  '1', oo.cell('B',1)
      assert_equal  '1', oo.cell('C',1)

      # assert_equal  1, oo.cell('A',2)
      # assert_equal  2, oo.cell('B',2)
      # assert_equal  1, oo.cell('C',2)
      # are stored as strings, not numbers

      assert_equal  1, oo.cell('A',2).to_i
      assert_equal  2, oo.cell('B',2).to_i
      assert_equal  1, oo.cell('C',2).to_i

      assert_equal  1, oo.cell('A',3)
      assert_equal  3, oo.cell('B',3)
      assert_equal  1, oo.cell('C',3)

      assert_equal  'A', oo.cell('A',4)
      assert_equal  'A', oo.cell('B',4)
      assert_equal  'A', oo.cell('C',4)

      # assert_equal  '0.01', oo.cell('A',5)
      # assert_equal  '0.01', oo.cell('B',5)
      # assert_equal  '0.01', oo.cell('C',5)
      #
      assert_equal  0.01, oo.cell('A',5)
      assert_equal  0.01, oo.cell('B',5)
      assert_equal  0.01, oo.cell('C',5)

      assert_equal 0.03, oo.cell('a',5)+oo.cell('b',5)+oo.cell('c',5)


      #   1.0

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
      assert_equal "0.01:percentage",oo.cell(5, 1).to_s+":"+oo.celltype(5, 1).to_s
      assert_equal "0.01:percentage",oo.cell(5, 2).to_s+":"+oo.celltype(5, 2).to_s
      assert_equal "0.01:percentage",oo.cell(5, 3).to_s+":"+oo.celltype(5, 3).to_s

      oo = Excel.new(File.join("test","simple_spreadsheet_from_italo.xls"))
      oo.default_sheet = oo.sheets.first

      assert_equal  '1', oo.cell('A',1)
      assert_equal  '1', oo.cell('B',1)
      assert_equal  '1', oo.cell('C',1)

      # assert_equal  1, oo.cell('A',2)
      # assert_equal  2, oo.cell('B',2)
      # assert_equal  1, oo.cell('C',2)
      # are stored as strings, not numbers

      assert_equal  1, oo.cell('A',2).to_i
      assert_equal  2, oo.cell('B',2).to_i
      assert_equal  1, oo.cell('C',2).to_i

      assert_equal  1, oo.cell('A',3)
      assert_equal  3, oo.cell('B',3)
      assert_equal  1, oo.cell('C',3)

      assert_equal  'A', oo.cell('A',4)
      assert_equal  'A', oo.cell('B',4)
      assert_equal  'A', oo.cell('C',4)

      # assert_equal  '0.01', oo.cell('A',5)
      # assert_equal  '0.01', oo.cell('B',5)
      # assert_equal  '0.01', oo.cell('C',5)
      #
      assert_equal  0.01, oo.cell('A',5)
      assert_equal  0.01, oo.cell('B',5)
      assert_equal  0.01, oo.cell('C',5)

      assert_equal 0.03, oo.cell('a',5)+oo.cell('b',5)+oo.cell('c',5)


      #   1.0

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
      #assert_equal "0.01:percentage",oo.cell(5, 1).to_s+":"+oo.celltype(5, 1).to_s
      #assert_equal "0.01:percentage",oo.cell(5, 2).to_s+":"+oo.celltype(5, 2).to_s
      #assert_equal "0.01:percentage",oo.cell(5, 3).to_s+":"+oo.celltype(5, 3).to_s
      # why do we get floats here? in the spreadsheet the cells were defined
      # to be percentage
      # TODO: should be fixed
      # the excel gem does not support the cell type 'percentage' these
      # cells are returned to be of the type float.
      assert_equal "0.01:float",oo.cell(5, 1).to_s+":"+oo.celltype(5, 1).to_s
      assert_equal "0.01:float",oo.cell(5, 2).to_s+":"+oo.celltype(5, 2).to_s
      assert_equal "0.01:float",oo.cell(5, 3).to_s+":"+oo.celltype(5, 3).to_s
    end
  end

  def test_formula_openoffice
    if OPENOFFICE
      oo = Openoffice.new(File.join("test","formula.ods"))
      oo.default_sheet = oo.sheets.first
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

  def test_formula_excel
    if defined? excel_supports_formulas
      if EXCEL
        oo = Excel.new(File.join("test","formula.xls"))
        oo.default_sheet = oo.sheets.first
        assert_equal 1, oo.cell('A',1)
        assert_equal 2, oo.cell('A',2)
        assert_equal 3, oo.cell('A',3)
        assert_equal 4, oo.cell('A',4)
        assert_equal 5, oo.cell('A',5)
        assert_equal 6, oo.cell('A',6)
        assert_equal :formula, oo.celltype('A',7)
        assert_equal 21, oo.cell('A',7)
        assert_equal " = [Sheet2.A1]", oo.formula('C',7)
        assert_nil oo.formula('A',6)
        assert_equal [[7, 1, " = SUM([.A1:.A6])"],
          [7, 2, " = SUM([.$A$1:.B6])"],
          [7, 3, " = [Sheet2.A1]"],
          [8, 2, " = SUM([.$A$1:.B7])"],
        ], oo.formulas

        # setting a cell
        oo.set('A',15, 41)
        assert_equal 41, oo.cell('A',15)
        oo.set('A',16, "41")
        assert_equal "41", oo.cell('A',16)
        oo.set('A',17, 42.5)
        assert_equal 42.5, oo.cell('A',17)

      end
    end
  end
  def test_formula_google
    if GOOGLE
      oo = Google.new(key_of("formula"))
      oo.default_sheet = oo.sheets.first
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
    end # GOOGLE
  end

  def test_formula_excelx
    if EXCELX
      oo = Excelx.new(File.join("test","formula.xlsx"))
      oo.default_sheet = oo.sheets.first
      assert_equal 1, oo.cell('A',1)
      assert_equal 2, oo.cell('A',2)
      assert_equal 3, oo.cell('A',3)
      assert_equal 4, oo.cell('A',4)
      assert_equal 5, oo.cell('A',5)
      assert_equal 6, oo.cell('A',6)
      assert_equal 21, oo.cell('A',7)
      assert_equal :formula, oo.celltype('A',7)
      after Date.new(9999,12,31) do
        #steht nicht in Datei, oder?
        #nein, diesen Bezug habe ich nur in der Openoffice-Datei
        assert_equal "=[Sheet2.A1]", oo.formula('C',7)
      end
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

  def test_borders_sheets_openoffice
    if OPENOFFICE
      oo = Openoffice.new(File.join("test","borders.ods"))
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

  def test_borders_sheets_excel
    if EXCEL
      oo = Excel.new(File.join("test","borders.xls"))
      oo.default_sheet = oo.sheets[1]
      assert_equal 6, oo.first_row
      assert_equal 11, oo.last_row
      assert_equal 4, oo.first_column
      assert_equal 8, oo.last_column

      oo.default_sheet = 1 # oo.sheets.first
      assert_equal 5, oo.first_row
      assert_equal 10, oo.last_row
      assert_equal 3, oo.first_column
      assert_equal 7, oo.last_column

      oo.default_sheet = 3 # oo.sheets[2]
      assert_equal 7, oo.first_row
      assert_equal 12, oo.last_row
      assert_equal 5, oo.first_column
      assert_equal 9, oo.last_column
    end
  end

  def test_borders_sheets_excelx
    if EXCELX
      oo = Excelx.new(File.join("test","borders.xlsx"))
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

  def test_borders_sheets_google
    if GOOGLE
      assert_nothing_raised(Timeout::Error) {
        Timeout::timeout(GLOBAL_TIMEOUT) do |timeout_length|
          oo = Google.new(key_of("borders"))
          oo.default_sheet = oo.sheets[0]
          assert_equal oo.sheets.first, oo.default_sheet
          assert_equal 5, oo.first_row
          oo.default_sheet = oo.sheets[1]
          assert_equal 'Sheet2', oo.default_sheet
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
      }
    end
  end

  def yaml_entry(row,col,type,value)
    "cell_#{row}_#{col}: \n  row: #{row} \n  col: #{col} \n  celltype: #{type} \n  value: #{value} \n"
  end

  def test_to_yaml
    if OPENOFFICE
      oo = Openoffice.new(File.join("test","numbers1.ods"))
      oo.default_sheet = oo.sheets.first
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
      #example: puts oo.to_yaml({}, 12,3)
      #example: puts oo.to_yaml({"probe" => "bodenproben_2007-06-30"}, 12,3)
    end
    if EXCEL
      oo = Excel.new(File.join("test","numbers1.xls"))
      oo.default_sheet = 1
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
    if EXCELX
      oo = Excelx.new(File.join("test","numbers1.xlsx"))
      oo.default_sheet = oo.sheets.first
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
    if GOOGLE
      oo = Google.new(key_of("numbers1"))
      oo.default_sheet = oo.sheets.first
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
      #example: puts oo.to_yaml({}, 12,3)
      #example: puts oo.to_yaml({"probe" => "bodenproben_2007-06-30"}, 12,3)
    end
  end

  if false
    def test_soap_server
      #threads = []
      #threads << Thread.new("serverthread") do
      fork do
        p "serverthread started"
        puts "in child, pid = #$$"
        puts `/usr/bin/ruby rooserver.rb`
        p "serverthread finished"
      end
      #threads << Thread.new("clientthread") do
      p "clientthread started"
      sleep 10
      proxy = SOAP::RPC::Driver.new("http://localhost:12321","spreadsheetserver")
      proxy.add_method('cell','row','col')
      proxy.add_method('officeversion')
      proxy.add_method('last_row')
      proxy.add_method('last_column')
      proxy.add_method('first_row')
      proxy.add_method('first_column')
      proxy.add_method('sheets')
      proxy.add_method('set_default_sheet','s')
      proxy.add_method('ferien_fuer_region', 'region')

      sheets = proxy.sheets
      p sheets
      proxy.set_default_sheet(sheets.first)

      assert_equal 1, proxy.first_row
      assert_equal 1, proxy.first_column
      assert_equal 187, proxy.last_row
      assert_equal 7, proxy.last_column
      assert_equal 42, proxy.cell('C',8)
      assert_equal 43, proxy.cell('F',12)
      assert_equal "1.0", proxy.officeversion
      p "clientthread finished"
      #end
      #threads.each {|t| t.join }
      puts "fertig"
      Process.kill("INT",pid)
      pid = Process.wait
      puts "child terminated, pid= #{pid}, status= #{$?.exitstatus}"
    end
  end # false

  def split_coord(s)
    letter = ""
    number = 0
    i = 0
    while i<s.length and "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz".include?(s[i,1])
      letter += s[i,1]
      i+=1
    end
    while i<s.length and "01234567890".include?(s[i,1])
      number = number*10 + s[i,1].to_i
      i+=1
    end
    if letter=="" or number==0
      raise ArgumentError
    end
    return letter,number
  end

  #def sum(s,expression)
  #  arg = expression.split(':')
  #  b,z = split_coord(arg[0])
  #  first_row = z
  #  first_col = Openoffice.letter_to_number(b)
  #  b,z = split_coord(arg[1])
  #  last_row = z
  #  last_col = Openoffice.letter_to_number(b)
  #  result = 0
  #  first_row.upto(last_row) {|row|
  #    first_col.upto(last_col) {|col|
  #      result = result + s.cell(row,col)
  #    }
  #  }
  #  result
  #end

  #def test_dsl
  #  s = Openoffice.new(File.join("test","numbers1.ods"))
  #  s.default_sheet = s.sheets.first
  #
  #    s.set 'a',1, 5
  #    s.set 'b',1, 3
  #    s.set 'c',1, 7
  #    s.set('a',2, s.cell('a',1)+s.cell('b',1))
  #    assert_equal 8, s.cell('a',2)
  #
  #    assert_equal 15, sum(s,'A1:C1')
  #  end

  #def test_create_spreadsheet1
  #  name = File.join('test','createdspreadsheet.ods')
  #  rm(name) if File.exists?(File.join('test','createdspreadsheet.ods'))
  #  # anlegen, falls noch nicht existierend
  #  s = Openoffice.new(name,true)
  #  assert File.exists?(name)
  #end

  #def test_create_spreadsheet2
  #  # anlegen, falls noch nicht existierend
  #  s = Openoffice.new(File.join("test","createdspreadsheet.ods"),true)
  #  s.set 'a',1,42
  #  s.set 'b',1,43
  #  s.set 'c',1,44
  #  s.save
  #
  #  #after Date.new(2007,7,3) do
  #  #  t = Openoffice.new(File.join("test","createdspreadsheet.ods"))
  #  #  assert_equal 42, t.cell(1,'a')
  #  #  assert_equal 43, t.cell('b',1)
  #  #  assert_equal 44, t.cell('c',3)
  #  #end
  #end

  def test_only_one_sheet
    if OPENOFFICE
      oo = Openoffice.new(File.join("test","only_one_sheet.ods"))
      # oo.default_sheet = oo.sheets.first
      assert_equal 42, oo.cell('B',4)
      assert_equal 43, oo.cell('C',4)
      assert_equal 44, oo.cell('D',4)
      oo.default_sheet = oo.sheets.first
      assert_equal 42, oo.cell('B',4)
      assert_equal 43, oo.cell('C',4)
      assert_equal 44, oo.cell('D',4)
    end
    if EXCEL
      oo = Excel.new(File.join("test","only_one_sheet.xls"))
      # oo.default_sheet = oo.sheets.first
      assert_equal 42, oo.cell('B',4)
      assert_equal 43, oo.cell('C',4)
      assert_equal 44, oo.cell('D',4)
      oo.default_sheet = oo.sheets.first
      assert_equal 42, oo.cell('B',4)
      assert_equal 43, oo.cell('C',4)
      assert_equal 44, oo.cell('D',4)
    end
    if EXCELX
      oo = Excelx.new(File.join("test","only_one_sheet.xlsx"))
      # oo.default_sheet = oo.sheets.first
      assert_equal 42, oo.cell('B',4)
      assert_equal 43, oo.cell('C',4)
      assert_equal 44, oo.cell('D',4)
      oo.default_sheet = oo.sheets.first
      assert_equal 42, oo.cell('B',4)
      assert_equal 43, oo.cell('C',4)
      assert_equal 44, oo.cell('D',4)
    end
    if GOOGLE
      oo = Google.new(key_of("only_one_sheet"))
      # oo.default_sheet = oo.sheets.first
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
        url = 'http://stiny-leonhard.de/bode-v1.xls.zip'
        excel = Excel.new(url, :zip)
        assert_equal 'ist "e" im Nenner von H(s)', excel.cell('b', 5)
        excel.remove_tmp # don't forget to remove the temporary files
      end
    end
  end

  #TODO: xlsx-Datei anpassen!
  def test_excelx_open_from_uri_and_zipped
    #TODO: gezippte xlsx Datei online zum Testen suchen
    after Date.new(2999,6,30) do
      if EXCELX
        if ONLINE
          url = 'http://stiny-leonhard.de/bode-v1.xlsx.zip'
          excel = Excelx.new(url, :zip)
          assert_equal 'ist "e" im Nenner von H(s)', excel.cell('b', 5)
          excel.remove_tmp # don't forget to remove the temporary files
        end
      end
    end
  end

  def test_openoffice_open_from_uri_and_zipped
    if OPENOFFICE
      if ONLINE
        url = 'http://spazioinwind.libero.it/s2/rata.ods.zip'
        sheet = Openoffice.new(url, :zip)
        #has been changed: assert_equal 'ist "e" im Nenner von H(s)', sheet.cell('b', 5)
        assert_in_delta 0.001, 505.14, sheet.cell('c', 33).to_f
        sheet.remove_tmp # don't forget to remove the temporary files
      end
    end
  end

  def SKIP_test_excel_zipped
    after Date.new(2009,1,10) do
      if EXCEL
        $log.level = Logger::DEBUG
        excel = Excel.new(File.join("test","bode-v1.xls.zip"), :zip)
        assert excel
        # muss Fehler bringen, weil kein default_sheet gesetzt wurde
        assert_raises (ArgumentError) {
          assert_equal 'ist "e" im Nenner von H(s)', excel.cell('b', 5)
        }
        $log.level = Logger::WARN
        excel.default_sheet = excel.sheets.first
        assert_equal 'ist "e" im Nenner von H(s)', excel.cell('b', 5)
        excel.remove_tmp # don't forget to remove the temporary files
      end
    end
  end

  def test_excelx_zipped
    # TODO: bode...xls bei Gelegenheit nach .xlsx konverieren lassen und zippen!
    if EXCELX
      after Date.new(2999,7,30) do
        # diese Datei gibt es noch nicht gezippt
        excel = Excelx.new(File.join("test","bode-v1.xlsx.zip"), :zip)
        assert excel
        assert_raises (ArgumentError) {
          assert_equal 'ist "e" im Nenner von H(s)', excel.cell('b', 5)
        }
        excel.default_sheet = excel.sheets.first
        assert_equal 'ist "e" im Nenner von H(s)', excel.cell('b', 5)
        excel.remove_tmp # don't forget to remove the temporary files
      end
    end
  end

  def test_openoffice_zipped
    if OPENOFFICE
      oo = Openoffice.new(File.join("test","bode-v1.ods.zip"), :zip)
      assert oo
      # muss Fehler bringen, weil kein default_sheet gesetzt wurde
      assert_raises (ArgumentError) {
        assert_equal 'ist "e" im Nenner von H(s)', oo.cell('b', 5)

      }
      oo.default_sheet = oo.sheets.first
      assert_equal 'ist "e" im Nenner von H(s)', oo.cell('b', 5)
      oo.remove_tmp # don't forget to remove the temporary files
    end
  end

  def test_bug_ric
    if OPENOFFICE
      oo = Openoffice.new(File.join("test","ric.ods"))
      oo.default_sheet = oo.sheets.first
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
      #assert_equal 2, oo.cell('f',1)
      #assert_equal 3, oo.cell('g',1)
      #assert_equal 4, oo.cell('h',1)
      #assert_equal 5, oo.cell('i',1)
      #assert_equal 6, oo.cell('j',1)
      #assert_equal 7, oo.cell('k',1)
      #assert_equal 8, oo.cell('l',1)
      #assert_equal 9, oo.cell('m',1)
      #assert_equal 10, oo.cell('n',1)
      #assert_equal 11, oo.cell('o',1)
      #assert_equal 12, oo.cell('p',1)
      #assert_equal 13, oo.cell('q',1)
      #assert_equal 14, oo.cell('r',1)
      #assert_equal 15, oo.cell('s',1)
      #assert_equal 16, oo.cell('t',1)
      #assert_equal 17, oo.cell('u',1)
      assert_equal 'J', oo.cell('v',1)
      assert_equal 'P', oo.cell('w',1)
      assert_equal 'B', oo.cell('x',1)
      assert_equal 'All', oo.cell('y',1)
      assert_equal 0, oo.cell('a',2)
      assert oo.empty?('b',2)
      assert oo.empty?('c',2)
      assert oo.empty?('d',2)

      #'e'.upto('s') {|letter|
      #  assert_equal 'B', oo.cell(letter,2)
      #}
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
    if OPENOFFICE
      oo = Openoffice.new(File.join("test","Bibelbund1.ods"))
      oo.default_sheet = oo.sheets.first
      assert_equal "Tagebuch des Sekret\303\244rs.    Letzte Tagung 15./16.11.75 Schweiz", oo.cell(45,'A')
    end
    #if EXCELX
    #  after Date.new(2008,6,1) do
    #    #Datei gibt es noch nicht
    #    oo = Excelx.new(File.join("test","Bibelbund1.xlsx"))
    #    oo.default_sheet = oo.sheets.first
    #    assert_equal "Tagebuch des Sekret\303\244rs.    Letzte Tagung 15./16.11.75 Schweiz", oo.cell(45,'A')
    #  end
    #end
  end

  def test_huge_document_to_csv_openoffice
    if LONG_RUN
      if OPENOFFICE
        assert_nothing_raised(Timeout::Error) {
          Timeout::timeout(GLOBAL_TIMEOUT) do |timeout_length|
            File.delete_if_exist("/tmp/Bibelbund.csv")
            oo = Openoffice.new(File.join("test","Bibelbund.ods"))
            oo.default_sheet = oo.sheets.first
            assert_equal "Tagebuch des Sekret\303\244rs.    Letzte Tagung 15./16.11.75 Schweiz", oo.cell(45,'A')
            assert_equal "Tagebuch des Sekret\303\244rs.  Nachrichten aus Chile", oo.cell(46,'A')
            assert_equal "Tagebuch aus Chile  Juli 1977", oo.cell(55,'A')
            assert oo.to_csv("/tmp/Bibelbund.csv")
            assert File.exists?("/tmp/Bibelbund.csv")
            assert_equal "", `diff test/Bibelbund.csv /tmp/Bibelbund.csv`
          end # Timeout
        } # nothing_raised
      end # OPENOFFICE
    end
  end

  def test_huge_document_to_csv_excel
    if LONG_RUN
      if EXCEL
        assert_nothing_raised(Timeout::Error) {
          Timeout::timeout(GLOBAL_TIMEOUT) do |timeout_length|
            File.delete_if_exist("/tmp/Bibelbund.csv")
            oo = Excel.new(File.join("test","Bibelbund.xls"))
            oo.default_sheet = oo.sheets.first
            assert oo.to_csv("/tmp/Bibelbund.csv")
            assert File.exists?("/tmp/Bibelbund.csv")
            assert_equal "", `diff test/Bibelbund.csv /tmp/Bibelbund.csv`
          end
        }
      end
    end # LONG_RUN
  end # def to_csv

  def test_huge_document_to_csv_excelx
    after Date.new(2008,8,27) do
      if LONG_RUN
        if EXCELX
          assert_nothing_raised(Timeout::Error) {
            Timeout::timeout(GLOBAL_TIMEOUT) do |timeout_length|
              File.delete_if_exist("/tmp/Bibelbund.csv")
              oo = Excelx.new(File.join("test","Bibelbund.xlsx"))
              oo.default_sheet = oo.sheets.first
              assert oo.to_csv("/tmp/Bibelbund.csv")
              assert File.exists?("/tmp/Bibelbund.csv")
              assert_equal "", `diff test/Bibelbund.csv /tmp/Bibelbund.csv`
            end
          }
        end
      end # LONG_RUN
    end
  end

  def test_huge_document_to_csv_google
    # maybe a better example... TODO:
    if GOOGLE and LONG_RUN
      assert_nothing_raised(Timeout::Error) {
        Timeout::timeout(GLOBAL_TIMEOUT) do |timeout_length|
          File.delete("/tmp/numbers1.csv") if File.exists?("/tmp/numbers1.csv")
          oo = Google.new(key_of('numbers1'))
          oo.default_sheet = oo.sheets.first
          #?? assert_equal "Tagebuch des Sekret\303\244rs.    Letzte Tagung 15./16.11.75 Schweiz", oo.cell(45,'A')
          #?? assert_equal "Tagebuch des Sekret\303\244rs.  Nachrichten aus Chile", oo.cell(46,'A')
          #?? assert_equal "Tagebuch aus Chile  Juli 1977", oo.cell(55,'A')
          assert oo.to_csv("/tmp/numbers1.csv")
          assert File.exists?("/tmp/numbers1.csv")
          assert_equal "", `diff test/numbers1.csv /tmp/numbers1.csv`
        end # Timeout
      } # nothing_raised
    end # GOOGLE
  end

  def test_to_csv_openoffice
    if OPENOFFICE
      #assert_nothing_raised(Timeout::Error) {
      Timeout::timeout(GLOBAL_TIMEOUT) do |timeout_length|
        File.delete_if_exist("/tmp/numbers1.csv")
        oo = Openoffice.new(File.join("test","numbers1.ods"))


        # bug?, 2008-01-15 from Troy Davis
        assert oo.to_csv("/tmp/numbers1.csv",oo.sheets.first)
        assert File.exists?("/tmp/numbers1.csv")
        assert_equal "", `diff test/numbers1.csv /tmp/numbers1.csv`

        oo.default_sheet = oo.sheets.first
        assert oo.to_csv("/tmp/numbers1.csv")
        assert File.exists?("/tmp/numbers1.csv")
        assert_equal "", `diff test/numbers1.csv /tmp/numbers1.csv`

      end # Timeout
      #} # nothing_raised
    end # OPENOFFICE
  end

  def test_to_csv_excel
    if EXCEL
      #assert_nothing_raised(Timeout::Error) {
      Timeout::timeout(GLOBAL_TIMEOUT) do |timeout_length|
        File.delete_if_exist("/tmp/numbers1_excel.csv")
        oo = Excel.new(File.join("test","numbers1.xls"))

        # bug?, 2008-01-15 from Troy Davis
        assert oo.to_csv("/tmp/numbers1_excel.csv",oo.sheets.first)
        assert File.exists?("/tmp/numbers1_excel.csv")
        assert_equal "", `diff test/numbers1_excel.csv /tmp/numbers1_excel.csv`
        oo.default_sheet = oo.sheets.first
        assert oo.to_csv("/tmp/numbers1_excel.csv")
        assert File.exists?("/tmp/numbers1_excel.csv")
        assert_equal "", `diff test/numbers1_excel.csv /tmp/numbers1_excel.csv`
      end

      #}
    end
  end

  def test_to_csv_excelx
    if EXCELX
      #assert_nothing_raised(Timeout::Error) {
      Timeout::timeout(GLOBAL_TIMEOUT) do |timeout_length|
        File.delete_if_exist("/tmp/numbers1.csv")
        oo = Excelx.new(File.join("test","numbers1.xlsx"))

        # bug?, 2008-01-15 from Troy Davis
        assert oo.to_csv("/tmp/numbers1.csv",oo.sheets.first)
        assert File.exists?("/tmp/numbers1.csv")
        assert_equal "", `diff test/numbers1.csv /tmp/numbers1.csv`
        oo.default_sheet = oo.sheets.first
        assert oo.to_csv("/tmp/numbers1.csv")
        assert File.exists?("/tmp/numbers1.csv")
        assert_equal "", `diff test/numbers1.csv /tmp/numbers1.csv`
      end

      #}
    end
  end

  def test_to_csv_google
    # maybe a better example... TODO:
    if GOOGLE
      #assert_nothing_raised(Timeout::Error) {
      Timeout::timeout(GLOBAL_TIMEOUT) do |timeout_length|
        File.delete_if_exist("/tmp/numbers1.csv") if File.exists?("/tmp/numbers1.csv")
        oo = Google.new(key_of('numbers1'))

        oo.default_sheet = oo.sheets.first
        assert oo.to_csv("/tmp/numbers1.csv")
        assert File.exists?("/tmp/numbers1.csv")
        assert_equal "", `diff test/numbers1.csv /tmp/numbers1.csv`

        # bug?, 2008-01-15 from Troy Davis
        assert oo.to_csv("/tmp/numbers1.csv",oo.sheets.first)
        assert File.exists?("/tmp/numbers1.csv")
        assert_equal "", `diff test/numbers1.csv /tmp/numbers1.csv`

      end # Timeout
      #} # nothing_raised
    end # GOOGLE
  end

  def test_bug_mehrere_datum
    if OPENOFFICE
      oo = Openoffice.new(File.join("test","numbers1.ods"))
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
    end # Openoffice

    if EXCEL
      oo = Excel.new(File.join("test","numbers1.xls"))
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
    end # Excel
    if EXCELX
      oo = Excelx.new(File.join("test","numbers1.xlsx"))
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
    end # Excelx
  end

  def test_multiple_sheets
    if OPENOFFICE
      oo = Openoffice.new(File.join("test","numbers1.ods"))
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
    end # OPENOFFICE


    if EXCEL
      $debug = true
      oo = Excel.new(File.join("test","numbers1.xls"))
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
      end # times
      $debug = false
    end # EXCEL
    if GOOGLE
      oo = Google.new(key_of("numbers1"))
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
    if EXCELX
      oo = Excelx.new(File.join("test","numbers1.xlsx"))
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


  def test_bug_empty_sheet_openoffice
    if OPENOFFICE
      oo = Openoffice.new(File.join("test","formula.ods"))
      oo.default_sheet = 'Sheet3' # is an empty sheet
      assert_nothing_raised(NoMethodError) {
        oo.to_csv(File.join("/","tmp","emptysheet.csv"))
      }
      assert_equal "", `cat /tmp/emptysheet.csv`
    end
  end

  def test_bug_empty_sheet_excelx
    if EXCELX
      oo = Excelx.new(File.join("test","formula.xlsx"))
      oo.default_sheet = 'Sheet3' # is an empty sheet
      assert_nothing_raised(NoMethodError) {
        oo.to_csv(File.join("/","tmp","emptysheet.csv"))
      }
      assert_equal "", `cat /tmp/emptysheet.csv`
    end
  end

  def test_find_by_row_huge_document_openoffice
    if LONG_RUN
      if OPENOFFICE
        Timeout::timeout(GLOBAL_TIMEOUT) do |timeout_length|
          oo = Openoffice.new(File.join("test","Bibelbund.ods"))
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

  def test_find_by_row_huge_document_excel
    if LONG_RUN
      if EXCEL
        Timeout::timeout(GLOBAL_TIMEOUT) do |timeout_length|
          oo = Excel.new(File.join("test","Bibelbund.xls"))
          oo.default_sheet = oo.sheets.first
          rec = oo.find 20
          assert rec
          #jetzt als Hash assert_equal "Brief aus dem Sekretariat", rec[0]
          assert_equal "Brief aus dem Sekretariat", rec[0]['TITEL']

          rec = oo.find 22
          assert rec
          # assert_equal "Brief aus dem Skretariat. Tagung in Amberg/Opf.",rec[0]
          assert_equal "Brief aus dem Skretariat. Tagung in Amberg/Opf.",rec[0]['TITEL']
        end
      end
    end
  end

  def test_find_by_row_huge_document_excelx
    if LONG_RUN
      if EXCEL
        Timeout::timeout(GLOBAL_TIMEOUT) do |timeout_length|
          oo = Excelx.new(File.join("test","Bibelbund.xlsx"))
          oo.default_sheet = oo.sheets.first
          rec = oo.find 20
          assert rec
          assert_equal "Brief aus dem Sekretariat", rec[0]['TITEL']

          rec = oo.find 22
          assert rec
          assert_equal "Brief aus dem Skretariat. Tagung in Amberg/Opf.",rec[0]['TITEL']
        end
      end
    end
  end

  def test_find_by_row_openoffice
    if OPENOFFICE
      oo = Openoffice.new(File.join("test","numbers1.ods"))
      oo.default_sheet = oo.sheets.first
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

  def test_find_by_row_excel
    if EXCEL
      oo = Excel.new(File.join("test","numbers1.xls"))
      oo.default_sheet = oo.sheets.first
      oo.header_line = nil
      rec = oo.find 16
      assert rec
      # keine Headerlines in diesem Beispiel definiert
      assert_equal "einundvierzig", rec[0]

      rec = oo.find 15
      assert rec
      assert_equal 41,rec[0]
    end
  end

  def test_find_by_row_excelx
    if EXCELX
      oo = Excelx.new(File.join("test","numbers1.xlsx"))
      oo.default_sheet = oo.sheets.first
      oo.header_line = nil
      rec = oo.find 16
      assert rec
      # keine Headerlines in diesem Beispiel definiert
      assert_equal "einundvierzig", rec[0]

      rec = oo.find 15
      assert rec
      assert_equal 41,rec[0]
    end
  end

  def test_find_by_row_google
    if GOOGLE
      oo = Google.new(key_of("numbers1"))
      oo.default_sheet = oo.sheets.first
      oo.header_line = nil
      rec = oo.find 16
      assert rec
      # keine Headerlines in diesem Beispiel definiert
      assert_equal "einundvierzig", rec[0]

      rec = oo.find 15
      assert rec
      assert_equal 41,rec[0]
    end
  end

  def test_find_by_row_huge_document_google
    if LONG_RUN
      if GOOGLE
        Timeout::timeout(GLOBAL_TIMEOUT) do |timeout_length|
          oo = Google.new(key_of("Bibelbund"))
          oo.default_sheet = oo.sheets.first
          rec = oo.find 20
          assert rec
          assert_equal "Brief aus dem Sekretariat", rec[0]

          rec = oo.find 22
          assert rec
          assert_equal "Brief aus dem Skretariat. Tagung in Amberg/Opf.",rec[0]
        end
      end
    end
  end

  def test_find_by_conditions_openoffice
    if LONG_RUN
      if OPENOFFICE
        assert_nothing_raised(Timeout::Error) {
          Timeout::timeout(GLOBAL_TIMEOUT) do |timeout_length|
            oo = Openoffice.new(File.join("test","Bibelbund.ods"))
            oo.default_sheet = oo.sheets.first
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

  def test_find_by_conditions_excel
    if LONG_RUN
      if EXCEL
        assert_nothing_raised(Timeout::Error) {
          Timeout::timeout(GLOBAL_TIMEOUT) do |timeout_length|
            oo = Excel.new(File.join("test","Bibelbund.xls"))
            oo.default_sheet = oo.sheets.first
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
          end # Timeout
        } # nothing_raised
      end
    end
  end

  #TODO: temporaerer Test
  def test_seiten_als_date
    oo = Excelx.new(File.join("test","Bibelbund.xlsx"))
    oo.default_sheet = oo.sheets.first
    assert_equal 'Bericht aus dem Sekretariat', oo.cell(13,1)
    assert_equal '1981-4', oo.cell(13,'D')
    assert_equal [:numeric_or_formula,"General"], oo.excelx_type(13,'E')
    assert_equal '428', oo.excelx_value(13,'E')
    assert_equal 428.0, oo.cell(13,'E')
  end

  def test_find_by_conditions_excelx
    if LONG_RUN
      if EXCELX
        assert_nothing_raised(Timeout::Error) {
          Timeout::timeout(GLOBAL_TIMEOUT) do |timeout_length|
            oo = Excelx.new(File.join("test","Bibelbund.xlsx"))
            oo.default_sheet = oo.sheets.first
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
          end # Timeout
        } # nothing_raised
      end
    end
  end

  def test_find_by_conditions_google
    if LONG_RUN
      if GOOGLE
        assert_nothing_raised(Timeout::Error) {
          Timeout::timeout(GLOBAL_TIMEOUT) do |timeout_length|
            oo = Google.new(key_of("Bibelbund"))
            oo.default_sheet = oo.sheets.first
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

  def test_column_openoffice
    after Date.new(2008,9,30) do

      expected = [1.0,5.0,nil,10.0,Date.new(1961,11,21),'tata',nil,nil,nil,nil,'thisisa11',41.0,nil,nil,41.0,'einundvierzig',nil,Date.new(2007,5,31)]
      if OPENOFFICE
        Timeout::timeout(GLOBAL_TIMEOUT) do |timeout_length|
          oo = Openoffice.new(File.join('test','numbers1.ods'))
          oo.default_sheet = oo.sheets.first
          assert_equal expected, oo.column(1)
          assert_equal expected, oo.column('a')
        end
      end
    end
  end

  def test_column_excel
    expected = [1.0,5.0,nil,10.0,Date.new(1961,11,21),'tata',nil,nil,nil,nil,'thisisa11',41.0,nil,nil,41.0,'einundvierzig',nil,Date.new(2007,5,31)]
    if EXCEL
      Timeout::timeout(GLOBAL_TIMEOUT) do |timeout_length|
        oo = Excel.new(File.join('test','numbers1.xls'))
        oo.default_sheet = oo.sheets.first
        assert_equal expected, oo.column(1)
        assert_equal expected, oo.column('a')
      end
    end
  end

  def test_column_excelx
    expected = [1.0,5.0,nil,10.0,Date.new(1961,11,21),'tata',nil,nil,nil,nil,'thisisa11',41.0,nil,nil,41.0,'einundvierzig',nil,Date.new(2007,5,31)]
    if EXCELX
      Timeout::timeout(GLOBAL_TIMEOUT) do |timeout_length|
        oo = Excelx.new(File.join('test','numbers1.xlsx'))
        oo.default_sheet = oo.sheets.first
        assert_equal expected, oo.column(1)
        assert_equal expected, oo.column('a')
      end
    end
  end

  def test_column_google
    expected = [1.0,5.0,nil,10.0,Date.new(1961,11,21),'tata',nil,nil,nil,nil,'thisisa11',41.0,nil,nil,41.0,'einundvierzig',nil,Date.new(2007,5,31)]
    if GOOGLE
      Timeout::timeout(GLOBAL_TIMEOUT) do |timeout_length|
        oo = Google.new(key_of('numbers1'))
        oo.default_sheet = oo.sheets.first
        assert_equal expected, oo.column(1)
        assert_equal expected, oo.column('a')
      end
    end
  end

  def test_column_huge_document_openoffice
    if LONG_RUN
      if OPENOFFICE
        assert_nothing_raised(Timeout::Error) {
          Timeout::timeout(GLOBAL_TIMEOUT) do |timeout_length|
            oo = Openoffice.new(File.join('test','Bibelbund.ods'))
            oo.default_sheet = oo.sheets.first
            assert_equal 3735, oo.column('a').size
            #assert_equal 499, oo.column('a').size
          end
        }
      end
    end
  end

  def test_column_huge_document_excel
    if LONG_RUN
      if EXCEL
        assert_nothing_raised(Timeout::Error) {
          Timeout::timeout(GLOBAL_TIMEOUT) do |timeout_length|
            oo = Excel.new(File.join('test','Bibelbund.xls'))
            oo.default_sheet = oo.sheets.first
            assert_equal 3735, oo.column('a').size
            #assert_equal 499, oo.column('a').size
          end
        }
      end
    end
  end

  def test_column_huge_document_excelx
    if LONG_RUN
      if EXCELX
        assert_nothing_raised(Timeout::Error) {
          Timeout::timeout(GLOBAL_TIMEOUT) do |timeout_length|
            oo = Excelx.new(File.join('test','Bibelbund.xlsx'))
            oo.default_sheet = oo.sheets.first
            assert_equal 3735, oo.column('a').size
            #assert_equal 499, oo.column('a').size
          end
        }
      end
    end
  end

  def test_column_huge_document_google
    if LONG_RUN
      if GOOGLE
        assert_nothing_raised(Timeout::Error) {
          Timeout::timeout(GLOBAL_TIMEOUT) do |timeout_length|
            #puts Time.now.to_s + "column Openoffice gestartet"
            oo = Google.new(key_of('Bibelbund'))
            oo.default_sheet = oo.sheets.first
            #assert_equal 3735, oo.column('a').size
            assert_equal 499, oo.column('a').size
            #puts Time.now.to_s + "column Openoffice beendet"
          end
        }
      end
    end
  end

  def test_simple_spreadsheet_find_by_condition_openoffice
    oo = Openoffice.new(File.join("test","simple_spreadsheet.ods"))
    oo.default_sheet = oo.sheets.first
    oo.header_line = 3
    erg = oo.find(:all, :conditions => {'Comment' => 'Task 1'})
    assert_equal Date.new(2007,05,07), erg[1]['Date']
    assert_equal 10.75       , erg[1]['Start time']
    assert_equal 12.50       , erg[1]['End time']
    assert_equal 0           , erg[1]['Pause']
    assert_equal 1.75        , erg[1]['Sum']
    assert_equal "Task 1"    , erg[1]['Comment']
  end

  def test_simple_spreadsheet_find_by_condition_excel
    if EXCEL
      $debug = true
      oo = Excel.new(File.join("test","simple_spreadsheet.xls"))
      oo.default_sheet = oo.sheets.first
      oo.header_line = 3
      erg = oo.find(:all, :conditions => {'Comment' => 'Task 1'})
      assert_equal Date.new(2007,05,07), erg[1]['Date']
      assert_equal 10.75       , erg[1]['Start time']
      assert_equal 12.50       , erg[1]['End time']
      assert_equal 0           , erg[1]['Pause']
      #cannot be tested because excel cannot return the result of formulas:
      #                    assert_equal 1.75        , erg[1]['Sum']
      assert_equal "Task 1"    , erg[1]['Comment']
      $debug = false
    end
  end

  def test_simple_spreadsheet_find_by_condition_excelx
    if EXCELX
      # die dezimalen Seiten bekomme ich seltsamerweise als Date
      oo = Excelx.new(File.join("test","simple_spreadsheet.xlsx"))
      oo.default_sheet = oo.sheets.first
      oo.header_line = 3
      erg = oo.find(:all, :conditions => {'Comment' => 'Task 1'})
      #expected =  { "Start time"=>10.75,
      #  "Pause"=>0.0,
      #  "Sum" => 1.75,
      #  "End time" => 12.5,
      #  "Pause" => 0.0,
      #  "Sum"=> 1.75,
      #  "Comment" => "Task 1",
      #  "Date" => Date.new(2007,5,7)}
      assert_equal Date.new(2007,5,7), erg[1]['Date']
      assert_equal 10.75,erg[1]['Start time']
      assert_equal 12.5, erg[1]['End time']
      assert_equal 0.0, erg[1]['Pause']
      assert_equal 1.75, erg[1]['Sum']
      assert_equal 'Task 1', erg[1]['Comment']

      #assert_equal expected, erg[1], erg[1]
      # hier bekomme ich den celltype :time zurueck
      # jetzt ist alles OK
      assert_equal Date.new(2007,05,07), erg[1]['Date']
      assert_equal 10.75       , erg[1]['Start time']
      assert_equal 12.50       , erg[1]['End time']
      assert_equal 0           , erg[1]['Pause']
      assert_equal 1.75        , erg[1]['Sum']
      assert_equal "Task 1"    , erg[1]['Comment']
    end
  end

  def test_simple_spreadsheet_find_by_condition_google
    if GOOGLE
      oo = Google.new(key_of("simple_spreadsheet"))
      oo.default_sheet = oo.sheets.first
      oo.header_line = 3
      erg = oo.find(:all, :conditions => {'Comment' => 'Task 1'})
      assert_equal Date.new(2007,05,07), erg[1]['Date']
      assert_equal 10.75       , erg[1]['Start time']
      assert_equal 12.50       , erg[1]['End time']
      assert_equal 0           , erg[1]['Pause']
      assert_kind_of Float, erg[1]['Sum']
      assert_equal 1.75        , erg[1]['Sum']
      assert_equal "Task 1"    , erg[1]['Comment']
    end
  end

  def DONT_test_false_encoding
    ex = Excel.new(File.join('test','false_encoding.xls'))
    ex.default_sheet = ex.sheets.first
    assert_equal "Sheet1", ex.sheets.first
    ex.first_row.upto(ex.last_row) do |row|
      ex.first_column.upto(ex.last_column) do |col|
        content = ex.cell(row,col)
        puts "#{row}/#{col}"
        #puts content if ! ex.empty?(row,col) or ex.formula?(row,col)
        if ex.formula?(row,col)
          #! ex.empty?(row,col)
          puts content
        end
      end
    end
  end

  def test_bug_false_borders_with_formulas
    if EXCEL
      after Date.new(2008,9,15) do
        ex = Excel.new(File.join('test','false_encoding.xls'))
        ex.default_sheet = ex.sheets.first
        #assert_equal 1, ex.first_row
=begin
  korrigiert auf Zeile 2. Zeile 1 enthaelt nur Formeln, die in parseexcel nicht
  ausgewertet werden koennen. D. h. der Nutzer hat keinen Vorteil davon, wenn
  er von Zeile 1 ab iterieren kann, da er auf die Formeln sowieso nicht zugreifen
  kann. Ideal waere aber noch eine Loesung, die auch diese Zeilen bei Excel
  als nichtleere Zeile liefert.
  TODO:
=end
        assert_equal 2, ex.first_row
        assert_equal 3, ex.last_row
        assert_equal 1, ex.first_column
        assert_equal 4, ex.last_column
      end
    end
  end

  def test_fe
    ex = Excel.new(File.join('test','false_encoding.xls'))
    ex.default_sheet = ex.sheets.first
    #DOES NOT WORK IN EXCEL FILES: assert_equal Date.new(2007,11,1), ex.cell('a',1)
    #DOES NOT WORK IN EXCEL FILES: assert_equal true, ex.formula?('a',1)
    #DOES NOT WORK IN EXCEL FILES: assert_equal '=TODAY()', ex.formula('a',1)

    #DOES NOT WORK IN EXCEL FILES: assert_equal Date.new(2008,2,9), ex.cell('B',1)
    #DOES NOT WORK IN EXCEL FILES: assert_equal true,               ex.formula?('B',1)
    #DOES NOT WORK IN EXCEL FILES: assert_equal "=A1+100",          ex.formula('B',1)

    #DOES NOT WORK IN EXCEL FILES: assert_equal Date.new(2008,2,9), ex.cell('C',1)
    #DOES NOT WORK IN EXCEL FILES: assert_equal true,               ex.formula?('C',1)
    #DOES NOT WORK IN EXCEL FILES: assert_equal "=C1",          ex.formula('C',1)

    assert_equal 'H1', ex.cell('A',2)
    assert_equal 'H2', ex.cell('B',2)
    assert_equal 'H3', ex.cell('C',2)
    assert_equal 'H4', ex.cell('D',2)
    assert_equal 'R1', ex.cell('A',3)
    assert_equal 'R2', ex.cell('B',3)
    assert_equal 'R3', ex.cell('C',3)
    assert_equal 'R4', ex.cell('D',3)
  end

  def test_excel_does_not_support_formulas
    if EXCEL
      ex = Excel.new(File.join('test','false_encoding.xls'))
      ex.default_sheet = ex.sheets.first
      assert_raise(RuntimeError) {
        void = ex.formula('a',1)
      }
      assert_raise(RuntimeError) {
        void = ex.formula?('a',1)
      }
      assert_raise(RuntimeError) {
        void = ex.formulas(ex.sheets.first)
      }
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
    if OPENOFFICE
      ext = ".ods"
      expected = sprintf(expected_templ,ext)
      oo = Openoffice.new(File.join("test","numbers1.ods"))
      assert_equal expected, oo.info
    end
    if EXCEL
      ext = ".xls"
      expected = sprintf(expected_templ,ext)
      oo = Excel.new(File.join("test","numbers1.xls"))
      assert_equal expected, oo.info
    end
    if EXCELX
      ext = ".xlsx"
      expected = sprintf(expected_templ,ext)
      oo = Excelx.new(File.join("test","numbers1.xlsx"))
      assert_equal expected, oo.info
    end
    if GOOGLE
      ext = ""
      expected = sprintf(expected_templ,ext)
      oo = Google.new(key_of("numbers1"))
      #$log.debug(expected)
      assert_equal expected.gsub(/numbers1/,key_of("numbers1")), oo.info
    end
  end

  def test_bug_excel_numbers1_sheet5_last_row
    if EXCEL
      oo = Excel.new(File.join("test","numbers1.xls"))
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
      after Date.new(2009,1,15) do
        assert_raise(IOError) {
          # oo = Google.new(key_of('testnichtvorhanden'+'Bibelbund.ods'))
          oo = Google.new('testnichtvorhanden')
        }
      end
    end
  end
  
  def test_bug_cell_no_default_sheet
    if GOOGLE
      oo = Google.new(key_of("numbers1"))
      assert_raise(ArgumentError) {
        # should complain about not set default-sheet
        #assert_equal 1.0, oo.cell('A',1)
        value = oo.cell('A',1)
        assert_equal "ganz rechts  gehts noch wetier", oo.cell('A',1,"Sheet3")
      }
    end
  end

  def test_write_google
    # write.me: http://spreadsheets.google.com/ccc?key=ptu6bbahNZpY0N0RrxQbWdw&hl=en_GB
    if GOOGLE
      oo = Google.new('ptu6bbahNZpY0N0RrxQbWdw')
      oo.default_sheet = oo.sheets.first
      oo.set_value(1,1,"hello from the tests")
      #oo.set_value(1,1,"sin(1)")
      assert_equal "hello from the tests", oo.cell(1,1)
    end
  end

  def test_bug_set_value_with_more_than_one_sheet_google
    # write.me: http://spreadsheets.google.com/ccc?key=ptu6bbahNZpY0N0RrxQbWdw&hl=en_GB
    if GOOGLE
      content1 = 'AAA'
      content2 = 'BBB'
      oo = Google.new('ptu6bbahNZpY0N0RrxQbWdw')
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
    if GOOGLE
      random_row = rand(10)+1
      random_column = rand(10)+1
      oo = Google.new('ptu6bbahNZpY0N0RrxQbWdw')
      oo.default_sheet = oo.sheets.first
      content1 = 'ABC'
      content2 = 'DEF'
      oo.set_value(random_row,random_column,content1,oo.sheets.first)
      oo.set_value(random_row,random_column,content2,oo.sheets[1])
      assert_equal content1, oo.cell(random_row,random_column,oo.sheets.first)
      assert_equal content2, oo.cell(random_row,random_column,oo.sheets[1])
    end
  end

  def test_set_value_for_non_existing_sheet_google
    if GOOGLE
      oo = Google.new('ptu6bbahNZpY0N0RrxQbWdw')
      assert_raise(RangeError) {
        #oo.default_sheet = "no_sheet"
        oo.set_value(1,1,"dummy","no_sheet")
      }
    end # GOOGLE
  end

  def test_bug_bbu_openoffice
    oo = Openoffice.new(File.join('test','bbu.ods'))
    assert_nothing_raised() {
      assert_equal "File: bbu.ods
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

  def test_bug_bbu_excel
    if EXCEL
      oo = Excel.new(File.join('test','bbu.xls'))
      assert_nothing_raised() {
        assert_equal "File: bbu.xls
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

  def test_bug_bbu_excelx
    if EXCELX
      oo = Excelx.new(File.join('test','bbu.xlsx'))
      assert_nothing_raised() {
        assert_equal "File: bbu.xlsx
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

  if false
    # there is no google spreadsheet for this test
    def test_bug_bbu_google
      oo = Excel.new(key_of('bbu'))
      assert_nothing_raised() {
        assert_equal "File: test/bbu.xls
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
  end # false

  def test_bug_time_nil_openoffice
    if OPENOFFICE
      oo = Openoffice.new(File.join("test","time-test.ods"))
      oo.default_sheet = oo.sheets.first
      assert_equal 12*3600+13*60+14, oo.cell('B',1) # 12:13:14 (secs since midnight)
      assert_equal :time, oo.celltype('B',1)
      assert_equal 15*3600+16*60, oo.cell('C',1) # 15:16    (secs since midnight)
      assert_equal :time, oo.celltype('C',1)

      assert_equal 23*3600, oo.cell('D',1) # 23:00    (secs since midnight)
      assert_equal :time, oo.celltype('D',1)
    end
  end

  def test_bug_time_nil_excel
    if EXCEL
      oo = Excel.new(File.join("test","time-test.xls"))
      oo.default_sheet = oo.sheets.first
      assert_equal 12*3600+13*60+14, oo.cell('B',1) # 12:13:14 (secs since midnight)
      assert_equal :time, oo.celltype('B',1)
      assert_equal 15*3600+16*60, oo.cell('C',1) # 15:16    (secs since midnight)
      assert_equal :time, oo.celltype('C',1)

      assert_equal 23*3600, oo.cell('D',1) # 23:00    (secs since midnight)
      assert_equal :time, oo.celltype('D',1)
    end
  end

  def test_bug_time_nil_excelx
    if EXCELX
      oo = Excelx.new(File.join("test","time-test.xlsx"))
      oo.default_sheet = oo.sheets.first
      assert_equal [:numeric_or_formula, "hh:mm:ss"],oo.excelx_type('b',1)
      assert_in_delta 0.50918981481481485, oo.excelx_value('b', 1), 0.000001
      assert_equal :time, oo.celltype('B',1)
      assert_equal 12*3600+13*60+14, oo.cell('B',1) # 12:13:14 (secs since midnight)

      assert_equal :time, oo.celltype('C',1)
      assert_equal 15*3600+16*60, oo.cell('C',1) # 15:16    (secs since midnight)

      assert_equal :time, oo.celltype('D',1)
      assert_equal 23*3600, oo.cell('D',1) # 23:00    (secs since midnight)
    end
  end

  def test_bug_time_nil_google
    if GOOGLE
      oo = Google.new(key_of("time-test"))
      oo.default_sheet = oo.sheets.first
      assert_equal 12*3600+13*60+14, oo.cell('B',1) # 12:13:14 (secs since midnight)
      assert_equal :time, oo.celltype('B',1)
      assert_equal 15*3600+16*60, oo.cell('C',1) # 15:16    (secs since midnight)
      assert_equal :time, oo.celltype('C',1)

      assert_equal 23*3600, oo.cell('D',1) # 23:00    (secs since midnight)
      assert_equal :time, oo.celltype('D',1)
    end
  end

  def test_date_time_to_csv_openoffice
    if OPENOFFICE
      File.delete_if_exist("/tmp/time-test.csv")
      oo = Openoffice.new(File.join("test","time-test.ods"))
      oo.default_sheet = oo.sheets.first
      assert oo.to_csv("/tmp/time-test.csv")
      assert File.exists?("/tmp/time-test.csv")
      assert_equal "", `diff test/time-test.csv /tmp/time-test.csv`
    end # OPENOFFICE
  end

  def test_date_time_to_csv_excel
    if EXCEL
      #ueberfluessige leere Zeilen werden am Ende noch angehaengt
      # last_row fehlerhaft?
      File.delete_if_exist("/tmp/time-test.csv")
      oo = Excel.new(File.join("test","time-test.xls"))
      oo.default_sheet = oo.sheets.first
      assert oo.to_csv("/tmp/time-test.csv")
      assert File.exists?("/tmp/time-test.csv")
      assert_equal "", `diff test/time-test.csv /tmp/time-test.csv`
    end # EXCEL
  end

  def test_date_time_to_csv_excelx
    if EXCELX
      #ueberfluessige leere Zeilen werden am Ende noch angehaengt
      # last_row fehlerhaft?
      File.delete_if_exist("/tmp/time-test.csv")
      oo = Excelx.new(File.join("test","time-test.xlsx"))
      oo.default_sheet = oo.sheets.first
      assert oo.to_csv("/tmp/time-test.csv")
      assert File.exists?("/tmp/time-test.csv")
      assert_equal "", `diff test/time-test.csv /tmp/time-test.csv`
    end # EXCELX
  end

  def test_date_time_to_csv_google
    if GOOGLE
      File.delete_if_exist("/tmp/time-test.csv")
      oo = Google.new(key_of("time-test"))
      oo.default_sheet = oo.sheets.first
      assert oo.to_csv("/tmp/time-test.csv")
      assert File.exists?("/tmp/time-test.csv")
      assert_equal "", `diff test/time-test.csv /tmp/time-test.csv`
    end # GOOGLE
  end

  def test_date_time_yaml_openoffice
    if OPENOFFICE
      expected =
        "--- \ncell_1_1: \n  row: 1 \n  col: 1 \n  celltype: string \n  value: Mittags: \ncell_1_2: \n  row: 1 \n  col: 2 \n  celltype: time \n  value: 12:13:14 \ncell_1_3: \n  row: 1 \n  col: 3 \n  celltype: time \n  value: 15:16:00 \ncell_1_4: \n  row: 1 \n  col: 4 \n  celltype: time \n  value: 23:00:00 \ncell_2_1: \n  row: 2 \n  col: 1 \n  celltype: date \n  value: 2007-11-21 \n"
      oo = Openoffice.new(File.join("test","time-test.ods"))
      oo.default_sheet = oo.sheets.first
      assert_equal expected, oo.to_yaml
    end
  end

  def test_date_time_yaml_excel
    if EXCEL
      expected =
        "--- \ncell_1_1: \n  row: 1 \n  col: 1 \n  celltype: string \n  value: Mittags: \ncell_1_2: \n  row: 1 \n  col: 2 \n  celltype: time \n  value: 12:13:14 \ncell_1_3: \n  row: 1 \n  col: 3 \n  celltype: time \n  value: 15:16:00 \ncell_1_4: \n  row: 1 \n  col: 4 \n  celltype: time \n  value: 23:00:00 \ncell_2_1: \n  row: 2 \n  col: 1 \n  celltype: date \n  value: 2007-11-21 \n"
      oo = Excel.new(File.join("test","time-test.xls"))
      oo.default_sheet = oo.sheets.first
      assert_equal expected, oo.to_yaml
    end
  end

  def test_date_time_yaml_excelx
    if EXCELX
      expected =
        "--- \ncell_1_1: \n  row: 1 \n  col: 1 \n  celltype: string \n  value: Mittags: \ncell_1_2: \n  row: 1 \n  col: 2 \n  celltype: time \n  value: 12:13:14 \ncell_1_3: \n  row: 1 \n  col: 3 \n  celltype: time \n  value: 15:16:00 \ncell_1_4: \n  row: 1 \n  col: 4 \n  celltype: time \n  value: 23:00:00 \ncell_2_1: \n  row: 2 \n  col: 1 \n  celltype: date \n  value: 2007-11-21 \n"
      oo = Excelx.new(File.join("test","time-test.xlsx"))
      oo.default_sheet = oo.sheets.first
      assert_equal expected, oo.to_yaml
    end
  end

  def test_date_time_yaml_google
    if GOOGLE
      expected =
        "--- \ncell_1_1: \n  row: 1 \n  col: 1 \n  celltype: string \n  value: Mittags: \ncell_1_2: \n  row: 1 \n  col: 2 \n  celltype: time \n  value: 12:13:14 \ncell_1_3: \n  row: 1 \n  col: 3 \n  celltype: time \n  value: 15:16:00 \ncell_1_4: \n  row: 1 \n  col: 4 \n  celltype: time \n  value: 23:00:00 \ncell_2_1: \n  row: 2 \n  col: 1 \n  celltype: date \n  value: 2007-11-21 \n"
      oo = Google.new(key_of("time-test"))
      oo.default_sheet = oo.sheets.first
      assert_equal expected, oo.to_yaml
    end
  end

  def test_no_remaining_tmp_files_openoffice
    if OPENOFFICE
      assert_raise(Zip::ZipError) { #TODO: besseres Fehlerkriterium bei
        # oo = Openoffice.new(File.join("test","no_spreadsheet_file.txt"))
        # es soll absichtlich ein Abbruch provoziert werden, deshalb :ignore
        oo = Openoffice.new(File.join("test","no_spreadsheet_file.txt"),
          false,
          :ignore)
      }
      a=Dir.glob("oo_*")
      assert_equal [], a
    end
  end

  def test_no_remaining_tmp_files_excel
    if EXCEL
      assert_raise(OLE::UnknownFormatError) {
        # oo = Excel.new(File.join("test","no_spreadsheet_file.txt"))
        # es soll absichtlich ein Abbruch provoziert werden, deshalb :ignore
        oo = Excel.new(File.join("test","no_spreadsheet_file.txt"),
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

        # oo = Excelx.new(File.join("test","no_spreadsheet_file.txt"))
        # es soll absichtlich ein Abbruch provoziert werden, deshalb :ignore
        oo = Excelx.new(File.join("test","no_spreadsheet_file.txt"),
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

  def do_test_xml(oo)
    assert_nothing_raised {oo.to_xml}
    sheetname = oo.sheets.first
    doc = REXML::Document.new(oo.to_xml)
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
          cell.text,
          cell.attributes['type'],
        ]
        assert_equal expected, result
        x += 1
      } # end of sheet
      sheetname = oo.sheets[oo.sheets.index(sheetname)+1]
    }
  end

  def test_to_xml_openoffice
    if OPENOFFICE
      oo = Openoffice.new(File.join('test','numbers1.ods'))
      do_test_xml(oo)
    end
  end

  def test_to_xml_excel
    if EXCEL
      oo = Excel.new(File.join('test','numbers1.xls'))
      do_test_xml(oo)
    end
  end

  def test_to_xml_excelx
    if EXCELX
      oo = Excelx.new(File.join('test','numbers1.xlsx'))
      do_test_xml(oo)
    end
  end

  def test_to_xml_google
    if GOOGLE
      oo = Google.new(key_of(File.join('test','numbers1.xlsx')))
      do_test_xml(oo)
    end
  end

  def SKIP_test_invalid_iconv_from_ms
    #TODO: does only run within a darwin-environment
    if   RUBY_PLATFORM.downcase =~ /darwin/
      assert_nothing_raised() {
        oo = Excel.new(File.join("test","ms.xls"))
      }
    end
  end

  def test_bug_row_column_fixnum_float
    if EXCEL
      ex = Excel.new(File.join('test','bug-row-column-fixnum-float.xls'))
      ex.default_sheet = ex.sheets.first
      assert_equal 42.5, ex.cell('b',2)
      assert_equal 43  , ex.cell('c',2)
      assert_equal ['hij',42.5, 43], ex.row(2)
      assert_equal ['def',42.5, 'nop'], ex.column(2)
    end
    
  end

  def test_bug_c2
    if EXCEL
      after Date.new(2009,1,6) do
        local_only do
          expected = ['Supermodel X','T6','Shaun White','Jeremy','Custom',
            'Warhol','Twin','Malolo','Supermodel','Air','Elite',
            'King','Dominant','Dominant Slick','Blunt','Clash',
            'Bullet','Tadashi Fuse','Jussi','Royale','S-Series',
            'Fish','Love','Feelgood ES','Feelgood','GTwin','Troop',
            'Lux','Stigma','Feather','Stria','Alpha','Feelgood ICS']
          result = []
          @e = Excel.new(File.join('test',"problem.xls"))
          @e.sheets[2..@e.sheets.length].each do |s|
            #(13..13).each do |s|
            @e.default_sheet = s
            name = @e.cell(2,'C')
            result << name
            #puts "#{name} (sheet: #{s})"
            #assert_equal "whatever (sheet: 13)",          "#{name} (sheet: #{s})"
          end
          assert_equal expected, result
        end
      end
    end
  end

  def test_bug_c2_parseexcel
    after Date.new(2009,1,10) do
      local_only do
        #-- this is OK
        @workbook = Spreadsheet::ParseExcel.parse(File.join('test',"problem.xls"))
        worksheet = @workbook.worksheet(11)
        skip = 0
        line = 1
        row = 2
        col = 3
        worksheet.each(skip) { |row_par|
          if line == row
            if row_par == nil
              raise "nil"
            end
            cell = row_par.at(col-1)
            assert cell, "cell should not be nil"
            assert_equal "Air", cell.to_s('utf-8')
          end
          line += 1
        }
        #-- worksheet 12 does not work
        @workbook = Spreadsheet::ParseExcel.parse(File.join('test',"problem.xls"))
        worksheet = @workbook.worksheet(12)
        skip = 0
        line = 1
        row = 2
        col = 3
        worksheet.each(skip) { |row_par|
          if line == row
            if row_par == nil
              raise "nil"
            end
            cell = row_par.at(col-1)
            assert cell, "cell should not be nil"
            assert_equal "Elite", cell.to_s('utf-8')
          end
          line += 1
        }
      end
    end
  end

  def test_bug_c2_excelx
    after Date.new(2008,9,15) do
      local_only do
        expected = ['Supermodel X','T6','Shaun White','Jeremy','Custom',
          'Warhol','Twin','Malolo','Supermodel','Air','Elite',
          'King','Dominant','Dominant Slick','Blunt','Clash',
          'Bullet','Tadashi Fuse','Jussi','Royale','S-Series',
          'Fish','Love','Feelgood ES','Feelgood','GTwin','Troop',
          'Lux','Stigma','Feather','Stria','Alpha','Feelgood ICS']
        result = []
        @e = Excelx.new(File.join('test',"problem.xlsx"))
        @e.sheets[2..@e.sheets.length].each do |s|
          @e.default_sheet = s
          #  assert_equal "A.",@e.cell('a',13)
          name = @e.cell(2,'C')
          result << name
          #puts "#{name} (sheet: #{s})"
          #assert_equal :string, @e.celltype('c',2)
          #assert_equal "Vapor (sheet: Vapor)", "#{name} (sheet: #{@e.sheets.first})"
          assert @e.cell(2,'c')
        end
        assert_equal expected, result

        @e = Excelx.new(File.join('test',"problem.xlsx"))
        #@e.sheets[2..@e.sheets.length].each do |s|
        (13..13).each do |s|
          @e.default_sheet = s
          name = @e.cell(2,'C')
          #puts "#{name} (sheet: #{s})"
          assert_equal "Elite (sheet: 13)",          "#{name} (sheet: #{s})"
        end
      end
    end
  end

  def test_compare_csv_excelx_excel
    if EXCELX
      after Date.new(2008,12,30) do
        # parseexcel bug
        local_only do
          s1 = Excel.new(File.join("test","problem.xls"))
          s2 = Excelx.new(File.join("test","problem.xlsx"))
          s1.sheets.each {|sh| #TODO:
            s1.default_sheet = sh
            s2.default_sheet = sh
            File.delete_if_exist("/tmp/problem.csv")
            File.delete_if_exist("/tmp/problemx.csv")
            assert s1.to_csv("/tmp/problem.csv")
            assert s2.to_csv("/tmp/problemx.csv")
            assert File.exists?("/tmp/problem.csv")
            assert File.exists?("/tmp/problemx.csv")
            assert_equal "", `diff /tmp/problem.csv /tmp/problemx.csv`, "Unterschied in Sheet #{sh} #{s1.sheets.index(sh)}"
          }
        end
      end
    end
  end

  def test_problemx_csv_imported
    after Date.new(2009,1,6) do
      if EXCEL
        local_only do
          # wieder eingelesene CSV-Datei aus obigem Test
          # muss identisch mit problem.xls sein
          # Importieren aus csv-Datei muss manuell gemacht werden
          ex = Excel.new(File.join("test","problem.xls"))
          cs = Excel.new(File.join("test","problemx_csv_imported.xls"))
          # nur das erste sheet betrachten
          ex.default_sheet = ex.sheets.first
          cs.default_sheet = cs.sheets.first
          ex.first_row.upto(ex.last_row) do |row|
            ex.first_column.upto(ex.last_column) do |col|
              assert_equal ex.cell(row,col), cs.cell(row,col), "cell #{row}/#{col} does not match '#{ex.cell(row,col)}' '#{cs.cell(row,col)}'"
              assert_equal ex.celltype(row,col), cs.celltype(row,col), "celltype #{row}/#{col} does not match"
              assert_equal ex.empty?(row,col), cs.empty?(row,col), "empty? #{row}/#{col} does not match"
              if defined? excel_supports_formulas
                assert_equal ex.formula?(row,col), cs.formula?(row,col), "formula? #{row}/#{col} does not match"
                assert_equal ex.formula(row,col), cs.formula(row,col), "formula #{row}/#{col} does not match"
              end
            end
          end
          cs.first_row.upto(cs.last_row) do |row|
            cs.first_column.upto(cs.last_column) do |col|
              assert_equal ex.cell(row,col), cs.cell(row,col), "cell #{row}/#{col} does not match '#{ex.cell(row,col)}' '#{cs.cell(row,col)}'"
              assert_equal ex.celltype(row,col), cs.celltype(row,col), "celltype #{row}/#{col} does not match"
              assert_equal ex.empty?(row,col), cs.empty?(row,col), "empty? #{row}/#{col} does not match"
              if defined? excel_supports_formulas
                assert_equal ex.formula?(row,col), cs.formula?(row,col), "formula? #{row}/#{col} does not match"
                assert_equal ex.formula(row,col), cs.formula(row,col), "formula #{row}/#{col} does not match"
              end
            end
          end
        end
      end
    end
  end

  def test_file_warning_default
    if OPENOFFICE
      assert_raises(TypeError) { oo = Openoffice.new(File.join("test","numbers1.xls")) }
      assert_raises(TypeError) { oo = Openoffice.new(File.join("test","numbers1.xlsx")) }
      assert_equal [], Dir.glob("oo_*")
    end
    if EXCEL
      assert_raises(TypeError) { oo = Excel.new(File.join("test","numbers1.ods")) }
      assert_raises(TypeError) { oo = Excel.new(File.join("test","numbers1.xlsx")) }
      assert_equal [], Dir.glob("oo_*")
    end
    if EXCELX
      assert_raises(TypeError) { oo = Excelx.new(File.join("test","numbers1.ods")) }
      assert_raises(TypeError) { oo = Excelx.new(File.join("test","numbers1.xls")) }
      assert_equal [], Dir.glob("oo_*")
    end
  end

  def test_file_warning_error
    if OPENOFFICE
      assert_raises(TypeError) { oo = Openoffice.new(File.join("test","numbers1.xls"),false,:error) }
      assert_raises(TypeError) { oo = Openoffice.new(File.join("test","numbers1.xlsx"),false,:error) }
      assert_equal [], Dir.glob("oo_*")
    end
    if EXCEL
      assert_raises(TypeError) { oo = Excel.new(File.join("test","numbers1.ods"),false,:error) }
      assert_raises(TypeError) { oo = Excel.new(File.join("test","numbers1.xlsx"),false,:error) }
      assert_equal [], Dir.glob("oo_*")
    end
    if EXCELX
      assert_raises(TypeError) { oo = Excelx.new(File.join("test","numbers1.ods"),false,:error) }
      assert_raises(TypeError) { oo = Excelx.new(File.join("test","numbers1.xls"),false,:error) }
      assert_equal [], Dir.glob("oo_*")
    end
  end

  def test_file_warning_warning
    if OPENOFFICE
      assert_nothing_raised(TypeError) {
        assert_raises(Zip::ZipError) {
          oo = Openoffice.new(File.join("test","numbers1.xls"),false, :warning)
        }
      }
      assert_nothing_raised(TypeError) {
        assert_raises(Errno::ENOENT) {
          oo = Openoffice.new(File.join("test","numbers1.xlsx"),false, :warning)
        }
      }
      assert_equal [], Dir.glob("oo_*")
    end
    if EXCEL
      assert_nothing_raised(TypeError) {
        assert_raises(OLE::UnknownFormatError) {
          oo = Excel.new(File.join("test","numbers1.ods"),false, :warning) }
      }
      assert_nothing_raised(TypeError) {
        assert_raises(OLE::UnknownFormatError) {
          oo = Excel.new(File.join("test","numbers1.xlsx"),false, :warning) }
      }
      assert_equal [], Dir.glob("oo_*")
    end
    if EXCELX
      assert_nothing_raised(TypeError) {
        assert_raises(Errno::ENOENT) {
          oo = Excelx.new(File.join("test","numbers1.ods"),false, :warning) }
      }
      assert_nothing_raised(TypeError) {
        assert_raises(Zip::ZipError) {
          oo = Excelx.new(File.join("test","numbers1.xls"),false, :warning) }
      }
      assert_equal [], Dir.glob("oo_*")
    end
  end

  def test_file_warning_ignore
    if OPENOFFICE
      assert_nothing_raised(TypeError) {
        assert_raises(Zip::ZipError) {
          oo = Openoffice.new(File.join("test","numbers1.xls"),false, :ignore) }
      }
      assert_nothing_raised(TypeError) {
        assert_raises(Errno::ENOENT) {
          oo = Openoffice.new(File.join("test","numbers1.xlsx"),false, :ignore) }
      }
      assert_equal [], Dir.glob("oo_*")
    end
    if EXCEL
      assert_nothing_raised(TypeError) {
        assert_raises(OLE::UnknownFormatError) {
          oo = Excel.new(File.join("test","numbers1.ods"),false, :ignore) }
      }
      assert_nothing_raised(TypeError) {
        assert_raises(OLE::UnknownFormatError) {oo = Excel.new(File.join("test","numbers1.xlsx"),false, :ignore) }}
      assert_equal [], Dir.glob("oo_*")
    end
    if EXCELX
      assert_nothing_raised(TypeError) {
        assert_raises(Errno::ENOENT) {
          oo = Excelx.new(File.join("test","numbers1.ods"),false, :ignore)
        }
      }
      assert_nothing_raised(TypeError) {
        assert_raises(Zip::ZipError) {
          oo = Excelx.new(File.join("test","numbers1.xls"),false, :ignore)
        }
      }
      assert_equal [], Dir.glob("oo_*")
    end
  end

  def SKIP_test_open_from_uri
    if ONLINE
      if OPENOFFICE
        assert_raises(RuntimeError) {
          oo = Openoffice.new("http://gibbsnichtdomainxxxxx.com/file.ods")
        }
      end
      if EXCEL
        assert_raises(RuntimeError) {
          oo = Excel.new("http://gibbsnichtdomainxxxxx.com/file.xls")
        }
      end
      if EXCELX
        assert_raises(RuntimeError) {
          oo = Excelx.new("http://gibbsnichtdomainxxxxx.com/file.xlsx")
        }
      end
    end
  end

  def test_bug_last_row_excel
    if EXCEL
      oo = Excel.new(File.join("test","time-test.xls"))
      oo.default_sheet = oo.sheets.first
      assert_equal 2, oo.last_row
    end
  end

  def test_bug_to_xml_with_empty_sheets_openoffice
    if OPENOFFICE
      oo = Openoffice.new(File.join("test","emptysheets.ods"))
      oo.sheets.each { |sheet|
        oo.default_sheet = sheet
        assert_equal nil, oo.first_row, "first_row not nil in sheet #{sheet}"
        assert_equal nil, oo.last_row, "last_row not nil in sheet #{sheet}"
        assert_equal nil, oo.first_column, "first_column not nil in sheet #{sheet}"
        assert_equal nil, oo.last_column, "last_column not nil in sheet #{sheet}"
        assert_equal nil, oo.first_row(sheet), "first_row not nil in sheet #{sheet}"
        assert_equal nil, oo.last_row(sheet), "last_row not nil in sheet #{sheet}"
        assert_equal nil, oo.first_column(sheet), "first_column not nil in sheet #{sheet}"
        assert_equal nil, oo.last_column(sheet), "last_column not nil in sheet #{sheet}"
      }
      assert_nothing_raised() {
        result = oo.to_xml
      }
    end
  end

  def test_bug_to_xml_with_empty_sheets_excel
    if EXCEL
      oo = Excel.new(File.join("test","emptysheets.xls"))
      oo.sheets.each { |sheet|
        oo.default_sheet = sheet
        assert_equal nil, oo.first_row, "first_row not nil in sheet #{sheet}"
        assert_equal nil, oo.last_row, "last_row not nil in sheet #{sheet}"
        assert_equal nil, oo.first_column, "first_column not nil in sheet #{sheet}"
        assert_equal nil, oo.last_column, "last_column not nil in sheet #{sheet}"
        assert_equal nil, oo.first_row(sheet), "first_row not nil in sheet #{sheet}"
        assert_equal nil, oo.last_row(sheet), "last_row not nil in sheet #{sheet}"
        assert_equal nil, oo.first_column(sheet), "first_column not nil in sheet #{sheet}"
        assert_equal nil, oo.last_column(sheet), "last_column not nil in sheet #{sheet}"
      }
      assert_nothing_raised() {
        result = oo.to_xml
      }
    end
  end

  def test_bug_to_xml_with_empty_sheets_excelx
    # kann ich nicht testen, da ich selbst keine .xlsx Files anlegen kann
    if EXCELX
      #   oo = Excelx.new(File.join("test","emptysheets.xlsx"))
      #  oo.sheets.each { |sheet|
      #    oo.default_sheet = sheet
      #    assert_equal nil, oo.first_row, "first_row not nil in sheet #{sheet}"
      #    assert_equal nil, oo.last_row, "last_row not nil in sheet #{sheet}"
      #    assert_equal nil, oo.first_column, "first_column not nil in sheet #{sheet}"
      #    assert_equal nil, oo.last_column, "last_column not nil in sheet #{sheet}"
      #    assert_equal nil, oo.first_row(sheet), "first_row not nil in sheet #{sheet}"
      #    assert_equal nil, oo.last_row(sheet), "last_row not nil in sheet #{sheet}"
      #    assert_equal nil, oo.first_column(sheet), "first_column not nil in sheet #{sheet}"
      #    assert_equal nil, oo.last_column(sheet), "last_column not nil in sheet #{sheet}"
      #  }
      #   assert_nothing_raised() {
      #     result = oo.to_xml
      # p result
      #   }
    end
  end

  def test_bug_simple_spreadsheet_time_bug
    # really a bug? are cells really of type time?
    # No! :float must be the correct type
    if EXCELX
      oo = Excelx.new(File.join("test","simple_spreadsheet.xlsx"))
      oo.default_sheet = oo.sheets.first
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


  def test_to_ascii_openoffice
    if OPENOFFICE
      after Date.new(9999,12,31) do
        oo = Openoffice.new(File.join("test","verysimple_spreadsheet.ods"))
        oo.default_sheet = oo.sheets.first
        expected="
  A    |   B   |  C   |
-------+-------+------|
      7|      8|     9|
-------+-------+------|
      4|      5|     6|
-------+-------+------|
      1|      2|     3|
----------------------/
        "
        assert_equal expected, oo.to_ascii
      end
    end
  end

  def test_simple2_excelx
    if EXCELX
      after Date.new(2008,8,2) do
        oo = Excelx.new(File.join("test","simple_spreadsheet.xlsx"))
        oo.default_sheet = oo.sheets.first
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
  end

  def test_possible_bug_snowboard_borders
    after Date.new(2008,12,15) do
      local_only do
        if EXCEL
          ex = Excel.new(File.join('test','problem.xls'))
          ex.default_sheet = ex.sheets.first
          assert_equal 2, ex.first_row
          assert_equal 30, ex.last_row
          assert_equal 'A', ex.first_column_as_letter
          assert_equal 'J', ex.last_column_as_letter
        end
        if EXCELX
          ex = Excelx.new(File.join('test','problem.xlsx'))
          ex.default_sheet = ex.sheets.first
          assert_equal 2, ex.first_row
          assert_equal 30, ex.last_row
          assert_equal 'A', ex.first_column_as_letter
          assert_equal 'J', ex.last_column_as_letter
        end
      end
    end
  end

  def common_possible_bug_snowboard_cells(ss)
    assert_equal "A.", ss.cell(13,'A'), ss.class
    assert_equal 147, ss.cell(13,'f'), ss.class
    assert_equal 152, ss.cell(13,'g'), ss.class
    assert_equal 156, ss.cell(13,'h'), ss.class
    assert_equal 158, ss.cell(13,'i'), ss.class
    assert_equal 160, ss.cell(13,'j'), ss.class
    assert_equal 164, ss.cell(13,'k'), ss.class
    assert_equal 168, ss.cell(13,'l'), ss.class
    assert_equal :string, ss.celltype(13,'m'), ss.class
    assert_equal "159W", ss.cell(13,'m'), ss.class
    assert_equal "164W", ss.cell(13,'n'), ss.class
    assert_equal "168W", ss.cell(13,'o'), ss.class
  end

  def test_possible_bug_snowboard_cells
    local_only do
      after Date.new(2009,1,6) do
        # warten auf Bugfix in parseexcel
        if EXCEL
          ex = Excel.new(File.join('test','problem.xls'))
          ex.default_sheet = 'Custom X'
          common_possible_bug_snowboard_cells(ex)
        end
      end
      if EXCELX
        ex = Excelx.new(File.join('test','problem.xlsx'))
        ex.default_sheet = 'Custom X'
        common_possible_bug_snowboard_cells(ex)
      end
    end
  end

  if EXCELX
    def test_possible_bug_2008_09_13
      local_only do
        # war nur in der 1.0.0 Release ein Fehler und sollte mit aktueller
        # Release nicht mehr auftreten.
=begin
−
	<sst count="46" uniqueCount="39">
−
	0<si>
<t>Bond</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	1<si>
<t>James</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	2<si>
<t>8659</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	3<si>
<t>12B</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	4<si>
<t>087692</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	5<si>
<t>Rowe</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	6<si>
<t>Karl</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	7<si>
<t>9128</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	8<si>
<t>79A</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	9<si>
<t>Benson</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	10<si>
<t>Cedric</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	11<si>
<t>Greenstreet</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	12<si>
<t>Jenny</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	13<si>
<t>Smith</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	14<si>
<t>Greame</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	15<si>
<t>Lucas</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	16<si>
<t>Ward</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	17<si>
<t>Lee</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	18<si>
<t>Bret</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	19<si>
<t>Warne</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	20<si>
<t>Shane</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	21<si>
<t>782</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	22<si>
<t>876</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	23<si>
<t>9901</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	24<si>
<t>1235</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	25<si>
<t>16547</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	26<si>
<t>7789</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	27<si>
<t>89</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	28<si>
<t>12A</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	29<si>
<t>19A</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	30<si>
<t>256</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	31<si>
<t>129B</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	32<si>
<t>11</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	33<si>
<t>Last Name</t>
</si>
−
	34<si>
<t>First Name</t>
</si>
−
35	<si>
<t>Middle Name</t>
</si>
−
	36<si>
<t>Resident ID</t>
</si>
−
	37<si>
<t>Room Number</t>
</si>
−
	38<si>
<t>Provider ID #</t>
</si>
</sst>
Hello Thomas,
How are you doing ? I am running into this strange issue with roo plugin (1.0.0). The attached
spreadsheet has all the cells formatted as "text", when I view in the Excel spreadsheet. But when it
get's into roo plugin (set_cell_values method - line 299), the values for the cells 1,1, 1,2, 1,3...1,6
show as 'date' instead of 'string'.
Because of this my parser is failing to get the proper values from the spreadsheet. Any ideas why
the formatting is getting set to the wrong value ?
Even stranger is if I save this file as ".XLS" and parse it the cells parse out fine as they are treated as
'string' instead of 'date'.
This attached file is the newer format of Microsoft Excel (.xlsx).

=end
        xx = Excelx.new(File.join('test','sample_file_2008-09-13.xlsx'))
        assert_equal 1, xx.sheets.size

        assert_equal 1, xx.first_row
        assert_equal 9, xx.last_row # 9 ist richtig. Es sind zwar 44 Zeilen definiert, aber der Rest hat keinen Inhalt
        assert_equal 1, xx.first_column
        assert_equal 6, xx.last_column
        assert_equal 'A', xx.first_column_as_letter
        assert_equal 'F', xx.last_column_as_letter

        assert_nothing_raised() {
          puts xx.info
        }
        p xx.cell(1,1)
        p xx.cell(1,2)
        p xx.cell(1,3)
        p xx.cell(1,4)
        p xx.cell(1,5)
        p xx.cell(1,6)
        xx.default_sheet = xx.sheets.first

        assert_equal 'Last Name', xx.cell('A',1)

        1.upto(6) do |col|
          assert_equal :string, xx.celltype(1,col)
        end
        #for col in (1..6)
        #  assert_equal "1234", xx.cell(1,col)
        #end
      end
    end
  end

  def do_datetime_tests(oo)
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
  
  def test_datetime_openoffice
    if OPENOFFICE
      oo = Openoffice.new(File.join("test","datetime.ods"))
      oo.default_sheet = oo.sheets.first
      do_datetime_tests(oo)
    end
  end

  def test_datetime_excel
    if EXCEL
      oo = Excel.new(File.join("test","datetime.xls"))
      oo.default_sheet = oo.sheets.first
      do_datetime_tests(oo)
    end
  end

  def test_datetime_excelx
    if EXCELX
      oo = Excelx.new(File.join("test","datetime.xlsx"))
      oo.default_sheet = oo.sheets.first
      do_datetime_tests(oo)
    end
  end

  def test_datetime_google
    if GOOGLE
      begin
        oo = Google.new(key_of('datetime'))
        oo.default_sheet = oo.sheets.first
        do_datetime_tests(oo)
      ensure
        $log.level = Logger::WARN
      end
    end
  end

  #-- bei diesen Test bekomme ich seltsamerweise einen Fehler can't allocate
  #-- memory innerhalb der zip-Routinen => erstmal deaktiviert
  def test_huge_table_timing_10_000_openoffice
    after Date.new(2009,1,1) do
      if OPENOFFICE
        if LONG_RUN
          assert_nothing_raised(Timeout::Error) {
            Timeout::timeout(3.minutes) do |timeout_length|
              oo = Openoffice.new("/home/tp/ruby-test/too-testing/speedtest_10000.ods")
              # process every cell
              sum = 0
              oo.sheets.each {|sheet|
                oo.default_sheet = sheet
                for row in oo.first_row..oo.last_row do
                  for col in oo.first_column..oo.last_column do
                    c = oo.cell(row,col)
                    sum += c.length if c
                  end
                end
                p sum
                assert sum > 0
              }
            end
          }
        end
      end
    end
  end

  def test_huge_table_timing_10_000_excel
    after Date.new(2009,1,1) do
      if EXCEL
        if LONG_RUN
          assert_nothing_raised(Timeout::Error) {
            Timeout::timeout(3.minutes) do |timeout_length|
              oo = Excel.new("/home/tp/ruby-test/too-testing/speedtest_10000.xls")
              # process every cell
              sum = 0
              oo.sheets.each {|sheet|
                oo.default_sheet = sheet
                for row in oo.first_row..oo.last_row do
                  for col in oo.first_column..oo.last_column do
                    c = oo.cell(row,col)
                    sum += c.length if c
                  end
                end
                p sum
                assert sum > 0
              }
            end
          }
        end
      end
    end
  end
  
  def test_huge_table_timing_10_000_google
    after Date.new(2009,1,1) do
      if GOOGLE
        if LONG_RUN
          assert_nothing_raised(Timeout::Error) {
            Timeout::timeout(3.minutes) do |timeout_length|
              oo = Excel.new(key_of("/home/tp/ruby-test/too-testing/speedtest_10000.xls"))
              # process every cell
              sum = 0
              oo.sheets.each {|sheet|
                oo.default_sheet = sheet
                for row in oo.first_row..oo.last_row do
                  for col in oo.first_column..oo.last_column do
                    c = oo.cell(row,col)
                    sum += c.length if c
                  end
                end
                p sum
                assert sum > 0
              }
            end
          }
        end
      end
    end
  end
  
  def test_huge_table_timing_10_000_excelx
    after Date.new(2009,1,1) do
      if EXCELX
        if LONG_RUN
          assert_nothing_raised(Timeout::Error) {
            Timeout::timeout(3.minutes) do |timeout_length|
              oo = Excelx.new("/home/tp/ruby-test/too-testing/speedtest_10000.xlsx")
              # process every cell
              sum = 0
              oo.sheets.each {|sheet|
                oo.default_sheet = sheet
                for row in oo.first_row..oo.last_row do
                  for col in oo.first_column..oo.last_column do
                    c = oo.cell(row,col)
                    sum += c.length if c
                  end
                end
                p sum
                assert sum > 0
              }
            end
          }
        end
      end
    end
  end
  
  # Eine Spreadsheetdatei wird nicht als Dateiname sondern direkt als Dokument
  # geoeffnettest_problemx_csv_imported
  def SKIP_test_from_stream_openoffice
    after Date.new(2009,1,6) do
      if OPENOFFICE
        filecontent = nil
        File.open(File.join("test","numbers1.ods")) do |f|
          filecontent = f.read
          p filecontent.class
          p filecontent.size
          #p filecontent
          assert filecontent.size > 0
          # #stream macht das gleiche wie #new liest abe aus Stream anstatt Datei
          oo = Openoffice.stream(filecontent)
        end
        #oo = Openoffice.open()
      end
    end
  end
  
  
  def SKIP_test_bug_encoding_exported_from_google
    if EXCEL
      xl = Excel.new(File.join("test","numbers1_from_google.xls"))
      xl.default_sheet = xl.sheets.first
      assert_equal 'test', xl.cell(2,'F')
    end
  end
  
  def test_cell_openoffice_html_escape
    if OPENOFFICE
      oo = Openoffice.new(File.join("test","html-escape.ods"))
      oo.default_sheet = oo.sheets.first
      assert_equal "'", oo.cell(1,1)
      assert_equal "&", oo.cell(2,1)
      assert_equal ">", oo.cell(3,1)
      assert_equal "<", oo.cell(4,1)
      assert_equal "`", oo.cell(5,1)
      # test_openoffice_zipped catches the &quot;
     end  
  end    
  
  def test_cell_excel_boolean
    if EXCEL
      oo = Excel.new(File.join("test","boolean.xls"))
      oo.default_sheet = oo.sheets.first
      assert_equal "TRUE", oo.cell(1,1)
      assert_equal "FALSE", oo.cell(2,1)
       end  
     if OPENOFFICE
       oo = Openoffice.new(File.join("test","boolean.ods"))
       oo.default_sheet = oo.sheets.first
       assert_equal "true", oo.cell(1,1)
       assert_equal "false", oo.cell(2,1)
     end  
     if EXCELX
       # TODO. need a source file to test with
     end
  end
  
end # class
