source 'https://rubygems.org'

gemspec

group :test do
  # additional testing libs
  gem 'shoulda'
  gem 'activesupport', '< 5.1'
  gem 'rspec', '>= 3.0.0'
  gem 'simplecov', '>= 0.9.0', require: false
  gem 'coveralls', require: false
  gem "minitest-reporters"
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
