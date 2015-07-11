require 'simplecov'
require 'roo'
require 'vcr'

require 'helpers'

RSpec.configure do |c|
  c.include Helpers
end

VCR.configure do |c|
  c.cassette_library_dir = 'spec/fixtures/vcr_cassettes'
  c.hook_into :webmock # or :fakeweb
end
