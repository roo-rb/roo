require 'jeweler'

Jeweler::Tasks.new do |gem|
  # gem is a Gem::Specification... see http://docs.rubygems.org/read/chapter/20 for more options
  gem.name = "roo"
  gem.summary = "Roo can access the contents of various spreadsheet files."
  gem.description = "Roo can access the contents of various spreadsheet files. It can handle\n* OpenOffice\n* Excel\n* Google spreadsheets\n* Excelx\n* LibreOffice\n* CSV"
  gem.email = "ruby.ruby.ruby.roo@gmail.com"
  gem.homepage = "http://github.com/Empact/roo"
  gem.authors = ['Thomas Preymesser', 'Hugh McGowan', 'Ben Woosley']

  gem.license = 'MIT'
  gem.required_ruby_version = '>= 1.9.0'

  gem.test_files = FileList["{spec,test}/**/*.*"]
end

require 'rake/testtask'
Rake::TestTask.new do |t|
  t.libs << "test"
  t.test_files = FileList['test/test*.rb']
  t.verbose = true
end
