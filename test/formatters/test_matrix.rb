require "test_helper"
require "matrix"

class TestRooFormatterMatrix < Minitest::Test
  def test_matrix
    expected_result = Matrix[
      [1.0, 2.0, 3.0],
      [4.0, 5.0, 6.0],
      [7.0, 8.0, 9.0]
    ]
    with_each_spreadsheet(name: "matrix", format: :openoffice) do |workbook|
      workbook.default_sheet = workbook.sheets.first
      assert_equal expected_result, workbook.to_matrix
    end
  end

  def test_matrix_selected_range
    expected_result = Matrix[
      [1.0, 2.0, 3.0],
      [4.0, 5.0, 6.0],
      [7.0, 8.0, 9.0]
    ]
    with_each_spreadsheet(name: "matrix", format: :openoffice) do |workbook|
      workbook.default_sheet = "Sheet2"
      assert_equal expected_result, workbook.to_matrix(3, 4, 5, 6)
    end
  end

  def test_matrix_all_nil
    expected_result = Matrix[
      [nil, nil, nil],
      [nil, nil, nil],
      [nil, nil, nil]
    ]
    with_each_spreadsheet(name: "matrix", format: :openoffice) do |workbook|
      workbook.default_sheet = "Sheet2"
      assert_equal expected_result, workbook.to_matrix(10, 10, 12, 12)
    end
  end

  def test_matrix_values_and_nil
    expected_result = Matrix[
      [1.0, nil, 3.0],
      [4.0, 5.0, 6.0],
      [7.0, 8.0, nil]
    ]
    with_each_spreadsheet(name: "matrix", format: :openoffice) do |workbook|
      workbook.default_sheet = "Sheet3"
      assert_equal expected_result, workbook.to_matrix(1, 1, 3, 3)
    end
  end

  def test_matrix_specifying_sheet
    expected_result = Matrix[
      [1.0, nil, 3.0],
      [4.0, 5.0, 6.0],
      [7.0, 8.0, nil]
    ]
    with_each_spreadsheet(name: "matrix", format: :openoffice) do |workbook|
      workbook.default_sheet = workbook.sheets.first
      assert_equal expected_result, workbook.to_matrix(nil, nil, nil, nil, "Sheet3")
    end
  end

  # #to_matrix of an empty sheet should return an empty matrix and not result in
  # an error message
  # 2011-06-25
  def test_bug_to_matrix_empty_sheet
    options = { name: "emptysheets", format: [:openoffice, :excelx] }
    with_each_spreadsheet(options) do |workbook|
      workbook.default_sheet = workbook.sheets.first
      workbook.to_matrix
      assert_equal(Matrix.empty(0, 0), workbook.to_matrix)
    end
  end
end
