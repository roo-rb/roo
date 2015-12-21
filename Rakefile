require 'bundler/gem_tasks'

require 'rake/testtask'
require 'rspec/core/rake_task'
require 'coveralls/rake/task'

# Test unit
Rake::TestTask.new do |t|
  t.libs << 'test'
  t.test_files = FileList['test/**/test*.rb']
  t.verbose = true
end

# RSpec
RSpec::Core::RakeTask.new(:spec)

# Coveralls
Coveralls::RakeTask.new

default_task = [:test, :spec]
default_task << 'coveralls:push' if ENV['TRAVIS']

task default: default_task
