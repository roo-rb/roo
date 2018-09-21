# frozen_string_literal: true

require "test_helper"

class TestRooExcelxCoordinate < Minitest::Test
  def row
    10
  end

  def column
    20
  end

  def coordinate
    Roo::Excelx::Coordinate.new(row, column)
  end

  def array
    [row, column]
  end

  def test_row
    assert_same  row, coordinate.row
  end

  def test_column
    assert_same  column, coordinate.column
  end

  def test_frozen?
    assert coordinate.frozen?
  end

  def test_equality
    hash = {}
    hash[coordinate] = true
    assert hash.key?(coordinate)
    assert hash.key?(array)
  end

  def test_expand
    r, c = coordinate
    assert_same row, r
    assert_same column, c
  end

  def test_aref
    assert_same row, coordinate[0]
    assert_same column, coordinate[1]
  end
end
