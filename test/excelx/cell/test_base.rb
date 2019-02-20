require 'test_helper'

class TestRooExcelxCellBase < Minitest::Test
  def base
    Roo::Excelx::Cell::Base
  end

  def value
    'Hello World'
  end

  def test_cell_type_is_base
    cell = base.new(value, nil, [], nil, nil, nil)
    assert_equal :base, cell.type
  end

  def test_cell_value
    cell_value = value
    cell = base.new(cell_value, nil, [], nil, nil, nil)
    assert_equal cell_value, cell.cell_value
  end

  def test_not_empty?
    cell = base.new(value, nil, [], nil, nil, nil)
    refute cell.empty?
  end

  def test_presence
    cell = base.new(value, nil, [], nil, nil, nil)
    assert_equal cell, cell.presence
  end

  def test_cell_type_is_formula
    formula = true
    cell = base.new(value, formula, [], nil, nil, nil)
    assert_equal :formula, cell.type
  end

  def test_formula?
    formula = true
    cell = base.new(value, formula, [], nil, nil, nil)
    assert cell.formula?
  end

  def test_cell_type_is_link
    link = 'http://example.com'
    cell = base.new(value, nil, [], nil, link, nil)
    assert_equal :link, cell.type
  end

  def test_link?
    link = 'http://example.com'
    cell = base.new(value, nil, [], nil, link, nil)
    assert cell.link?
  end

  def test_link_value
    link = 'http://example.com'
    cell = base.new(value, nil, [], nil, link, nil)
    assert_equal value, cell.value
  end

  def test_link_value_href
    link = 'http://example.com'
    cell = base.new(value, nil, [], nil, link, nil)
    assert_equal link, cell.value.href
  end
end
