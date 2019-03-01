require "test_helper"

class TestImageIndex < Minitest::Test
  def test_sheets
    with_each_spreadsheet(name: "kangaroos", format: [:excelx]) do |oo|
      assert_equal 1, oo.sheet_for(0).images.size
      assert_equal 0, oo.sheet_for(1).images.size
      assert_equal 1, oo.sheet_for(2).images.size
    end
  end
end
