require 'test_helper'

class TestRooFormulaToCsv < Minitest::Test
  def cell_to_csv(row, col)
    Roo::Spreadsheet.open(
      File.join(TESTDIR, 'formula_cell_types.xlsx')
    ).send('cell_to_csv', row, col, 'Sheet1')
  end

  def test_true_class
    assert_equal 'true', cell_to_csv(1, 1)
  end

  def test_false_class
    assert_equal 'false', cell_to_csv(2, 1)
  end

  def test_date_class
    assert_equal '2017-01-01', cell_to_csv(3, 1)
  end
end
