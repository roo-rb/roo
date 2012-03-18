require 'test/unit'
require File.dirname(__FILE__) + '/../lib/roo'

# helper method
def after(d)
  yield if DateTime.now > d
end
  
# helper method
def before(d)
  yield if DateTime.now <= d
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
def diff(fn1,fn2)
  result = ''
  f1 = File.open(fn1)
  f2 = File.open(fn2)
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
  f1.close
  f2.close
  result
end
