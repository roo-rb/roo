require 'simplecov'
require 'roo'
require 'helpers'

RSpec.configure do |c|
  c.include Helpers
  c.color = true
  c.formatter = :documentation
end
