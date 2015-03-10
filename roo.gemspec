# encoding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'roo/version'

Gem::Specification.new do |spec|
  spec.name          = 'roo'
  spec.version       = Roo::VERSION
  spec.authors       = ['Thomas Preymesser', 'Hugh McGowan', 'Ben Woosley']
  spec.email         = ['ruby.ruby.ruby.roo@gmail.com']
  spec.summary       = 'Roo can access the contents of various spreadsheet files.'
  spec.description   = "Roo can access the contents of various spreadsheet files. It can handle the following formats\n* OpenOffice\n*Excelx\n* LibreOffice\n* CSV"
  spec.homepage      = 'http://github.com/roo-rb/roo'
  spec.license       = 'MIT'

  spec.files         = `git ls-files -z`.split("\x0")
  spec.executables   = spec.files.grep(%r{^bin/}) { |f| File.basename(f) }
  spec.test_files    = spec.files.grep(%r{^(test|spec|features)/})
  spec.require_paths = ['lib']

  spec.add_dependency 'nokogiri', '~> 1.5'
  spec.add_dependency 'rubyzip', '~> 1.1', '>= 1.1.7'

  spec.add_development_dependency 'rake', '~> 10.1'
  spec.add_development_dependency 'minitest', '~> 5.4', '>= 5.4.3'
end
