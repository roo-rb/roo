require "test_helper"

class TestRooExcelxElapsedTime < Minitest::Test
  def test_elapsed_formats_over_a_day_are_seconds
    workbook = Roo::Excelx.new(File.join(TESTDIR, "elapsed_time.xlsx"))

    assert_equal 93_600, workbook.cell(1, 1)
    assert_equal :time, workbook.celltype(1, 1)
    assert_equal 129_600, workbook.cell(2, 1)
    assert_equal 172_800, workbook.cell(3, 1)
    assert_equal 43_200, workbook.cell(4, 1)
    assert_equal :datetime, workbook.celltype(5, 1)
  end
end
