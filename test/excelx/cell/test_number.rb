require 'test_helper'

class TestRooExcelxCellNumber < Minitest::Test
  def number
    Roo::Excelx::Cell::Number
  end

  def test_float
    cell = Roo::Excelx::Cell::Number.new '42.1', nil, ['General'], nil, nil, nil
    assert_kind_of(Float, cell.value)
  end

  def test_integer
    cell = Roo::Excelx::Cell::Number.new '42', nil, ['0'], nil, nil, nil
    assert_kind_of(Integer, cell.value)
  end

  def test_scientific_notation
    cell = Roo::Excelx::Cell::Number.new '1.2E-3', nil, ['0'], nil, nil, nil
    assert_kind_of(Float, cell.value)
  end

  def test_simple_scientific_notation
    cell = Roo::Excelx::Cell::Number.new '1E-3', nil, ['0'], nil, nil, nil
    assert_kind_of(Float, cell.value)
  end

  def test_very_simple_scientific_notation
    cell = Roo::Excelx::Cell::Number.new '1e6', nil, ['0'], nil, nil, nil
    assert_kind_of(Float, cell.value)
  end

  def test_percent
    cell = Roo::Excelx::Cell::Number.new '42.1', nil, ['0.00%'], nil, nil, nil
    assert_kind_of(Float, cell.value)
  end

  def test_formats_with_negative_numbers
    [
      ['#,##0 ;(#,##0)', '(1,042)'],
      ['#,##0 ;[Red](#,##0)', '[Red](1,042)'],
      ['#,##0.00;(#,##0.00)', '(1,042.00)'],
      ['#,##0.00;[Red](#,##0.00)', '[Red](1,042.00)']
    ].each do |style_format, result|
      cell = Roo::Excelx::Cell::Number.new '-1042', nil, [style_format], nil, nil, nil
      assert_equal result, cell.formatted_value, "Style=#{style_format}"
    end
  end

  def test_numbers_with_cell_errors
    Roo::Excelx::ERROR_VALUES.each do |error|
      cell = Roo::Excelx::Cell::Number.new error, nil, ['General'], nil, nil, nil
      assert_equal error, cell.value
      assert_equal error, cell.formatted_value
    end
  end

  def test_formats
    [
      ['General', '1042'],
      ['GENERAL', '1042'],
      ['0', '1042'],
      ['000000', '001042'],
      ['0.00', '1042.00'],
      ['0.0000', '1042.0000'],
      ['0.000000000', '1042.000000000'],
      ['#,##0', '1,042'],
      ['#,##0.00', '1,042.00'],
      ['0%', '104200%'],
      ['0.00%', '104200.00%'],
      ['0.00E+00', '1.04E+03'],
      ['#,##0 ;(#,##0)', '1,042'],
      ['#,##0 ;[Red](#,##0)', '1,042'],
      ['#,##0.00;(#,##0.00)', '1,042.00'],
      ['#,##0.00;[Red](#,##0.00)', '1,042.00'],
      ['##0.0E+0', '1.0E+03'],
      ['@', '1042']
    ].each do |style_format, result|
      cell = Roo::Excelx::Cell::Number.new '1042', nil, [style_format], nil, nil, nil
      assert_equal result, cell.formatted_value, "Style=#{style_format}"
    end
  end
end
