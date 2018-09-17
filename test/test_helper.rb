# encoding: utf-8
require "simplecov"
require "tmpdir"
require "fileutils"
require "minitest/autorun"
require "shoulda"
require "timeout"
require "logger"
require "date"

# require gem files
require "roo"
require "minitest/reporters"
if ENV["USE_REPORTERS"]
  Minitest::Reporters.use!(
    [
      Minitest::Reporters::DefaultReporter.new,
      Minitest::Reporters::SpecReporter.new
    ]
  )
end

TESTDIR = File.join(File.dirname(__FILE__), "files")
ROO_FORMATS = [
  :excelx,
  :excelxm,
  :openoffice,
  :libreoffice
]

require "helpers/test_accessing_files"
require "helpers/test_comments"
require "helpers/test_formulas"
require "helpers/test_labels"
require "helpers/test_sheets"
require "helpers/test_styles"


# very simple diff implementation
# output is an empty string if the files are equal
# otherwise differences a printen (not compatible to
# the diff command)
def file_diff(fn1,fn2)
  result = ""
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
          result << ">#{line2}\n"
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

def local_server(port)
  raise ArgumentError unless port.to_i > 0
  "http://0.0.0.0:#{port}"
end

def start_local_server(filename, port = nil)
  require "rack"
  content_type = filename.split(".").last
  port ||= TEST_RACK_PORT

  web_server = Proc.new do |env|
    [
      "200",
      { "Content-Type" => content_type },
      [File.read("#{TESTDIR}/#{filename}")]
    ]
  end

  t = Thread.new { Rack::Handler::WEBrick.run web_server, Host: "0.0.0.0", Port: port , Logger: WEBrick::BasicLog.new(nil,1) }
  # give the app a chance to startup
  sleep(0.2)

  yield
ensure
  t.kill
end

# call a block of code for each spreadsheet type
# and yield a reference to the roo object
def with_each_spreadsheet(options)
  if options[:format]
    formats = Array(options[:format])
    invalid_formats = formats - ROO_FORMATS
    unless invalid_formats.empty?
      raise "invalid spreadsheet types: #{invalid_formats.join(', ')}"
    end
  else
    formats = ROO_FORMATS
  end
  formats.each do |format|
    begin
      yield Roo::Spreadsheet.open(File.join(TESTDIR,
        fixture_filename(options[:name], format)))
    rescue => e
      raise e, "#{e.message} for #{format}", e.backtrace unless options[:ignore_errors]
    end
  end
end

def get_extension(oo)
  case oo
  when Roo::OpenOffice
    ".ods"
  when Roo::Excelx
    ".xlsx"
  end
end

def fixture_filename(name, format)
  case format
  when :excelx
    "#{name}.xlsx"
  when :excelxm
    "#{name}.xlsm"
  when :openoffice, :libreoffice
    "#{name}.ods"
  else
    raise ArgumentError, "unexpected format #{format}"
  end
end

def skip_long_test
  msg = "This is very slow, test use `LONG_RUN=true bundle exec rake` to run it"
  skip(msg) unless ENV["LONG_RUN"]
end

def skip_jruby_incompatible_test
  msg = "This test uses a feature incompatible with JRuby"
  skip(msg) if defined?(JRUBY_VERSION)
end

def with_timezone(new_tz)
  if new_tz
    begin
      prev_tz, ENV['TZ'] = ENV['TZ'], new_tz
      yield
    ensure
      ENV['TZ'] = prev_tz
    end
  else
    yield
  end
end
