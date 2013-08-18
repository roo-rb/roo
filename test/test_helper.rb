# require deps
require 'tmpdir'
require 'fileutils'
require 'test/unit'
require 'shoulda'
require 'fileutils'
require 'timeout'
require 'logger'
require 'date'
require 'webmock/test_unit'

# require gem files
require File.dirname(__FILE__) + '/../lib/roo'

TESTDIR =  File.join(File.dirname(__FILE__), 'files')

LOG_DIR = File.join(File.dirname(__FILE__),'../log')
FileUtils.mkdir_p(LOG_DIR)

LOG_FILE = File.join(LOG_DIR,'roo_test.log')
$log = Logger.new(LOG_FILE)

#$log.level = Logger::WARN
$log.level = Logger::DEBUG

DISPLAY_LOG = false

# very simple diff implementation
# output is an empty string if the files are equal
# otherwise differences a printen (not compatible to
# the diff command)
def file_diff(fn1,fn2)
  result = ''
  File.open(fn1) do |f1|
    File.open(fn2) do |f2|
      while f1.eof? == false and f2.eof? == false
        line1 = f1.gets.chomp
        line2 = f2.gets.chomp
        result << "<#{line1}\n>#{line2}\n" if line1 != line2
      end
      if f1.eof? == false
        while f1.eof? == false
          line1 = f1.gets
          result << "<#{line1}\n"
        end
      end
      if f2.eof? == false
        while f2.eof? == false
          line2 = f2.gets
          result ">#{line2}\n"
        end
      end
    end
  end
  result
end

class File
  def File.delete_if_exist(filename)
    if File.exist?(filename)
      File.delete(filename)
    end
  end
end

class Test::Unit::TestCase
  def key_of(spreadsheetname)
    {
      #'formula' => 'rt4Pw1WmjxFtyfrqqy94wPw',
      'formula' => 'o10837434939102457526.3022866619437760118',
      #"write.me" => 'r6m7HFlUOwst0RTUTuhQ0Ow',
      "write.me" => '0AkCuGANLc3jFcHR1NmJiYWhOWnBZME4wUnJ4UWJXZHc',
      #'numbers1' => "rYraCzjxTtkxw1NxHJgDU8Q",
      'numbers1' => 'o10837434939102457526.4784396906364855777',
      #'borders' => "r_nLYMft6uWg_PT9Rc2urXw",
      'borders' => "o10837434939102457526.664868920231926255",
      #'simple_spreadsheet' => "r3aMMCBCA153TmU_wyIaxfw",
      'simple_spreadsheet' => "ptu6bbahNZpYe-L1vEBmgGA",
      'testnichtvorhandenBibelbund.ods' => "invalidkeyforanyspreadsheet", # !!! intentionally false key
      #"only_one_sheet" => "rqRtkcPJ97nhQ0m9ksDw2rA",
      "only_one_sheet" => "o10837434939102457526.762705759906130135",
      #'time-test' => 'r2XfDBJMrLPjmuLrPQQrEYw',
      'time-test' => 'ptu6bbahNZpYBMhk01UfXSg',
      #'datetime' => "r2kQpXWr6xOSUpw9MyXavYg",
      'datetime' => "ptu6bbahNZpYQEtZwzL_dZQ",
      'whitespace' => "rZyQaoFebVGeHKzjG6e9gRQ",
      'matrix' => '0AkCuGANLc3jFdHY3cWtYUkM4bVdadjZ5VGpfTzFEUEE',
    # 'numbers1' => "o10837434939102457526.4784396906364855777",
    # 'borders' => "o10837434939102457526.664868920231926255",
    # 'simple_spreadsheet' => "ptu6bbahNZpYe-L1vEBmgGA",
    # 'testnichtvorhandenBibelbund.ods' => "invalidkeyforanyspreadsheet", # !!! intentionally false key
    # "only_one_sheet" => "o10837434939102457526.762705759906130135",
    # "write.me" => 'ptu6bbahNZpY0N0RrxQbWdw&hl',
    # 'formula' => 'o10837434939102457526.3022866619437760118',
    # 'time-test' => 'ptu6bbahNZpYBMhk01UfXSg',
    # 'datetime' => "ptu6bbahNZpYQEtZwzL_dZQ",
    }.fetch(spreadsheetname)
  rescue KeyError
    raise "unknown spreadsheetname: #{spreadsheetname}"
  end

  def yaml_entry(row,col,type,value)
    "cell_#{row}_#{col}: \n  row: #{row} \n  col: #{col} \n  celltype: #{type} \n  value: #{value} \n"
  end
end
