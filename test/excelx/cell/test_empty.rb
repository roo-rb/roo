require 'test_helper'

class TestRooExcelxCellEmpty < Minitest::Test
  def empty
    Roo::Excelx::Cell::Empty
  end

  def test_empty?
    cell = empty.new(nil)
    assert_same  true, cell.empty?
  end

  def test_nil_presence
    cell = empty.new(nil)
    assert_nil cell.presence
  end

end
