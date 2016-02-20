require 'test_helper'

class TestRooExcelxCellDate < Minitest::Test
  def date_cell
    Roo::Excelx::Cell::Date
  end

  def base_date
    ::Date.new(1899, 12, 30)
  end

  def base_date_1904
    ::Date.new(1904, 01, 01)
  end

  def test_handles_1904_base_date
    cell = date_cell.new('41791', nil, [:numeric_or_formula, 'mm-dd-yy'], 6, nil, base_date_1904, nil)
    assert_equal ::Date.new(2018, 06, 02), cell.value
  end

  def test_formatted_value
    cell = date_cell.new('41791', nil, [:numeric_or_formula, 'mm-dd-yy'], 6, nil, base_date, nil)
    assert_equal '06-01-14', cell.formatted_value

    cell = date_cell.new('41791', nil, [:numeric_or_formula, 'yyyy-mm-dd'], 6, nil, base_date, nil)
    assert_equal '2014-06-01', cell.formatted_value
  end

  def test_value_is_date
    cell = date_cell.new('41791', nil, [:numeric_or_formula, 'mm-dd-yy'], 6, nil, base_date, nil)
    assert_kind_of ::Date, cell.value
  end

  def test_value
    cell = date_cell.new('41791', nil, [:numeric_or_formula, 'mm-dd-yy'], 6, nil, base_date, nil)
    assert_equal ::Date.new(2014, 06, 01), cell.value
  end
end
