require "bundler/gem_tasks"

require 'rake/testtask'
require 'rspec/core/rake_task'
require 'coveralls/rake/task'

# Test unit
Rake::TestTask.new do |t|
  t.libs << 'test'
  t.test_files = FileList['test/test*.rb']
  t.verbose = true
end

# RSpec
RSpec::Core::RakeTask.new(:spec)

# Coveralls
Coveralls::RakeTask.new

task test_with_coveralls: [:test, :spec, 'coveralls:push']

task default: [:test, :spec]
