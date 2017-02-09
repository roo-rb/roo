require 'simplecov'
require 'roo'
require 'helpers'

RSpec.configure do |c|
  c.include Helpers
  c.color = true
  c.formatter = :documentation if ENV["USE_REPORTERS"]
end
