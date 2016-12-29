require 'simplecov'
# require deps
require 'tmpdir'
require 'fileutils'
require 'minitest/autorun'
require 'shoulda'
require 'fileutils'
require 'timeout'
require 'logger'
require 'date'

# require gem files
require 'roo'

TESTDIR = File.join(File.dirname(__FILE__), 'files')
TEST_RACK_PORT = (ENV["ROO_TEST_PORT"] || 5000).to_i
TEST_URL= "http://0.0.0.0:#{TEST_RACK_PORT}"

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

def start_local_server(filename)
  require "rack"
  content_type = filename.split(".").last

  web_server = Proc.new do |env|
    [
      "200",
      { "Content-Type" => content_type },
      [File.read("#{TESTDIR}/#{filename}")]
    ]
  end

  t = Thread.new { Rack::Handler::WEBrick.run web_server, Port: TEST_RACK_PORT, Logger: WEBrick::BasicLog.new(nil,1) }
  # give the app a chance to startup
  sleep(0.5)

  yield
ensure
  t.kill
end
