require 'spec_helper'

RSpec.describe Roo::Excelx::Cell::Time do
  it "should set the cell value to the correct number of seconds" do
    value = 0.05513888888888888 # '1:19:24'
    excelx_type = [:numeric_or_formula, "h:mm:ss"]
    base_timestamp = Date.new(1899, 12, 30).to_time.to_i
    time_cell = Roo::Excelx::Cell::Time.new(value, nil, excelx_type, 1, nil, base_timestamp, nil)
    expect(time_cell.value).to eq(1*60*60 + 19*60 + 24) # '1:19:24' in seconds
    # use case from https://github.com/roo-rb/roo/issues/310
    value = 0.523761574074074   # '12:34:13' in seconds
    time_cell = Roo::Excelx::Cell::Time.new(value, nil, excelx_type, 1, nil, base_timestamp, nil)
    expect(time_cell.value).to eq(12*60*60 + 34*60 + 13) # 12:34:13 in seconds
  end
end