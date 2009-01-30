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