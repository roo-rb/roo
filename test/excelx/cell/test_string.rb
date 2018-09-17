require 'test_helper'

class TestRooExcelxCellString < Minitest::Test
  def string
    Roo::Excelx::Cell::String
  end


  def test_formatted_value
    cell = string.new '1', nil, nil, nil, nil
    assert_equal '1', cell.formatted_value
  end

  def test_to_s
    cell = string.new '0', nil, nil, nil, nil
    assert_equal '0', cell.to_s
  end

  def test_cell_value
    cell = string.new '1', nil, nil, nil, nil
    assert_equal '1', cell.cell_value
  end

  def test_value
    cell = string.new '0', nil, nil, nil, nil
    assert_equal '0', cell.value
  end

  def test_not_empty?
    cell = string.new '1', nil, nil, nil, nil
    assert_equal false, cell.empty?
  end

  def test_empty?
    cell = string.new '', nil, nil, nil, nil
    assert_equal true, cell.empty?
  end

  def test_presence
    cell = string.new '1', nil, nil, nil, nil
    assert_equal cell, cell.presence
  end

  def test_nil_presence
    cell = string.new '', nil, nil, nil, nil
    assert_nil cell.presence
  end
end
