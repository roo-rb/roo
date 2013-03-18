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
DB_LOG = false

if DB_LOG
  require 'activerecord'

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


class Roo::Csv
  remove_method :cell_postprocessing
  def cell_postprocessing(row,col,value)
    if row==1 and col==1
      return value.to_f
    end
    if row==1 and col==2
      return value.to_s
    end
    return value
  end
end

# helper method
def local_only
  if ENV["roo_local"] == "thomas-p"
    yield
  end
end

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

# :nodoc
class Fixnum
  def minutes
    self * 60
  end
end

class Test::Unit::TestCase
  def key_of(spreadsheetname)
    return {
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

  def yaml_entry(row,col,type,value)
    "cell_#{row}_#{col}: \n  row: #{row} \n  col: #{col} \n  celltype: #{type} \n  value: #{value} \n"
  end

  if DB_LOG
    if ! (defined?(@connected) and @connected)
      activerecord_connect
    else
      @connected = true
    end
  end
  # alias unlogged_run run
  # def run(result, &block)
  #   t1 = Time.now
  #   if DISPLAY_LOG
  #       v1,v2,_ = RUBY_VERSION.split('.')
  #       if v1.to_i > 1 or
  #         (v1.to_i == 1 and v2.to_i > 8)
  #         # Ruby 1.9.x
  #       print "RUNNING #{self.class} #{self.__name__} \t#{Time.now.to_s}"
  #       else
  #         # Ruby < 1.9.x
  #       print "RUNNING #{self.class} #{@method_name} \t#{Time.now.to_s}"
  #       end
  #     STDOUT.flush
  #   end
  #   unlogged_run result, &block
  #   t2 = Time.now
  #   if DISPLAY_LOG
  #     puts "\t#{t2-t1} seconds"
  #   end
  #   if DB_LOG
  #     Testrun.create(
  #       :class_name => self.class.to_s,
  #       :test_name => @method_name,
  #       :start => t1,
  #       :duration => t2-t1
  #     )
  #   end
  # end
end
