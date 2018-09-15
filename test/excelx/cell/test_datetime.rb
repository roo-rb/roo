require 'test_helper'

class TestRooExcelxCellDateTime < Minitest::Test
  def test_cell_value_is_datetime
    cell = datetime.new('30000.323212', nil, ['mm-dd-yy'], nil, nil, base_timestamp, nil)
    assert_kind_of ::DateTime, cell.value
  end

  def test_cell_type_is_datetime
    cell = datetime.new('30000.323212', nil, [], nil, nil, base_timestamp, nil)
    assert_equal :datetime, cell.type
  end

  def test_standard_formatted_value
    [
      ['mm-dd-yy', '01-25-15'],
      ['d-mmm-yy', '25-JAN-15'],
      ['d-mmm ', '25-JAN'],
      ['mmm-yy', 'JAN-15'],
      ['m/d/yy h:mm', '1/25/15 8:15']
    ].each do |format, formatted_value|
      cell = datetime.new '42029.34375', nil, [format], nil, nil, base_timestamp, nil
      assert_equal formatted_value, cell.formatted_value
    end
  end

  def test_custom_formatted_value
    [
      ['yyyy/mm/dd hh:mm:ss', '2015/01/25 08:15:00'],
      ['h:mm:ss000 mm/yy', '8:15:00000 01/15'],
      ['mmm yyy', '2015-01-25 08:15:00']
    ].each do |format, formatted_value|
      cell = datetime.new '42029.34375', nil, [format], nil, nil, base_timestamp, nil
      assert_equal formatted_value, cell.formatted_value
    end
  end

  def datetime
    Roo::Excelx::Cell::DateTime
  end

  def base_timestamp
    DateTime.new(1899, 12, 30).to_time.to_i
  end
end
