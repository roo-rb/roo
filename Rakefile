$:.unshift('lib')
 
require 'rake/testtask'

begin
  require 'jeweler'
  Jeweler::Tasks.new do |s|
    s.name               = 'roo'
    s.rubyforge_project  = 'roo'
    s.platform           = Gem::Platform::RUBY
    s.email              = 'thopre@gmail.com' 
    s.homepage           = "http://roo.rubyforge.org"
    s.summary            = "roo"
    s.description        = "roo can access the contents of OpenOffice-, Excel- or Google-Spreadsheets"
    s.authors            = ['Thomas Preymesser']
    s.files              =  FileList[ "{lib,test}/**/*"]
    s.has_rdoc = true
    s.extra_rdoc_files = ["README.txt", "History.txt"]
    s.rdoc_options = ["--main","README.txt"]
    s.add_dependency "spreadsheet", [">= 0.6.3.1"]
    s.add_dependency "rubyzip", [">= 0.9.1"]
    s.add_dependency "hpricot", [">= 0.5"]
    s.add_dependency "GData", [">= 0.0.3"]
    
  end
rescue LoadError
  puts "Jeweler not available. Install it with: sudo gem install technicalpickles-jeweler -s http://gems.github.com"
end

Rake::TestTask.new do |t|
  t.libs << "test"
  t.test_files = FileList['test/test_roo.rb']
  t.verbose = true
end


task :default => :test