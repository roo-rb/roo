require 'simplecov'
require 'roo'
require 'helpers'

RSpec.configure do |c|
  c.include Helpers
  c.color = true
  c.formatter = :documentation if ENV["USE_REPORTERS"]
  original_stderr = $stderr
  original_stdout = $stdout
  c.before(:all) do
    # Redirect stderr and stdout
    $stderr = File.open(File::NULL, "w")
    $stdout = File.open(File::NULL, "w")
  end
  c.after(:all) do
    $stderr = original_stderr
    $stdout = original_stdout
  end
end
