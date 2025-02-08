# encoding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'roo/version'

Gem::Specification.new do |spec|
  spec.name                   = 'roo'
  spec.version                = Roo::VERSION
  spec.authors                = ['Thomas Preymesser', 'Hugh McGowan', 'Ben Woosley', 'Oleksandr Simonov', 'Steven Daniels', 'Anmol Chopra']
  spec.email                  = ['ruby.ruby.ruby.roo@gmail.com', 'oleksandr@simonov.me']
  spec.summary                = 'Roo can access the contents of various spreadsheet files.'
  spec.description            = "Roo can access the contents of various spreadsheet files. It can handle\n* OpenOffice\n* Excelx\n* LibreOffice\n* CSV"
  spec.homepage               = 'https://github.com/roo-rb/roo'
  spec.license                = 'MIT'

  spec.files                  = `git ls-files -z`.split("\x0")
  spec.files.reject! { |fn| fn.include?('test/files') }
  spec.require_paths          = ['lib']

  if defined?(RUBY_ENGINE) && RUBY_ENGINE == 'jruby'
    spec.required_ruby_version  = ">= 2.6.0"
  else
    spec.required_ruby_version  = ">= 2.7.0"
  end

  if RUBY_VERSION >= '3.4.0'
    spec.add_dependency 'base64', '~> 0.2'
    spec.add_dependency 'csv', '~> 3'
  end
  spec.add_dependency 'nokogiri', '~> 1'
  spec.add_dependency 'rubyzip', '>= 1.3.0', '< 3.0.0'

  spec.add_development_dependency 'rake'
  spec.add_development_dependency 'minitest', '~> 5.4', '>= 5.4.3'
  spec.add_development_dependency 'rack', '~> 1.6', '< 2.0.0'
  if RUBY_VERSION >= '3.0.0'
    spec.add_development_dependency 'matrix'
  end
end
