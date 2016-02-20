require 'test_helper'

class TestRooExcelxCellNumber < Minitest::Test
  def boolean
    Roo::Excelx::Cell::Boolean
  end

  def test_formatted_value
    cell = boolean.new '1', nil, nil, nil, nil
    assert_equal 'TRUE', cell.formatted_value

    cell = boolean.new '0', nil, nil, nil, nil
    assert_equal 'FALSE', cell.formatted_value
  end

  def test_to_s
    cell = boolean.new '1', nil, nil, nil, nil
    assert_equal 'TRUE', cell.to_s

    cell = boolean.new '0', nil, nil, nil, nil
    assert_equal 'FALSE', cell.to_s
  end

  def test_cell_value
    cell = boolean.new '1', nil, nil, nil, nil
    assert_equal '1', cell.cell_value
  end

  def test_value
    cell = boolean.new '1', nil, nil, nil, nil
    assert_equal true, cell.value

    cell = boolean.new '0', nil, nil, nil, nil
    assert_equal false, cell.value
  end
end
