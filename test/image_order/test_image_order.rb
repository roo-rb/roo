require "test_helper"
require 'pry'

class TestRooCSV < Minitest::Test
  def test_sheets
    ex = Roo::Excelx.new(File.expand_path("../kangaroos.xlsx", __FILE__).to_s)
    assert_equal 1, ex.sheet_for(0).images.size
    assert_equal 0, ex.sheet_for(1).images.size
    assert_equal 1, ex.sheet_for(2).images.size
  end
end
