source 'https://rubygems.org'

gemspec

gem 'rubocop'
gem 'rubocop-performance', require: false

group :test do
  # additional testing libs
  gem 'shoulda'
  gem 'activesupport'
  gem 'rspec', '>= 3.0.0'
  gem 'simplecov', '>= 0.9.0', require: false
  gem 'coveralls', require: false
  gem "minitest-reporters"
  gem 'webrick' 
end

group :local_development do
  gem 'terminal-notifier-guard', require: false if RUBY_PLATFORM.downcase.include?('darwin')
  gem 'guard-rspec', '>= 4.3.1', require: false
  gem 'guard-minitest', require: false
  gem 'guard-bundler', require: false
  gem 'guard-rubocop', require: false
  gem "rb-readline"
  gem 'pry'
end
