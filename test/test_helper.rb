require 'test/unit'
require File.dirname(__FILE__) + '/../lib/roo'
         
# helper method
def local_only
  if ENV["roo_local"] == "thomas-p"
    yield
  end
end

TMP_PREFIX = "#{Roo::GenericSpreadsheet::TEMP_PREFIX}*"
def assert_no_temp_files_left_over
  prev = Dir.glob(TMP_PREFIX)
  yield
  now = Dir.glob(TMP_PREFIX)
  assert (now-prev).empty?, "temporary directory not removed"
end

# very simple diff implementation
# output is an empty string if the files are equal
# otherwise differences a printen (not compatible to
# the diff command)
def diff(fn1,fn2)
  result = ''
  File.open(fn1) do |f1|
    File.open(fn2) do |f2|
      while f1.eof? == false and f2.eof? == false
        line1 = f1.gets
        line2 = f2.gets
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
