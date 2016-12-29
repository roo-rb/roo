require "test_helper"

class TestRooExcelx < Minitest::Test
  def test_download_uri_with_invalid_host
    assert_raises(RuntimeError) do
      Roo::Excelx.new("http://example.com/file.xlsx")
    end
  end

  def test_download_uri_with_query_string
    file = filename("simple_spreadsheet")
    url = "#{TEST_URL}/#{file}?query-param=value"

    start_local_server(file) do
      spreadsheet = roo_class.new(url)
      spreadsheet.default_sheet = spreadsheet.sheets.first
      assert_equal "Task 1", spreadsheet.cell("f", 4)
    end
  end

  def test_should_raise_file_not_found_error
    assert_raises(IOError) do
      Roo::Excelx.new(File.join("testnichtvorhanden", "Bibelbund.xlsx"))
    end
  end

  def test_file_warning_default
    assert_raises(TypeError) { Roo::Excelx.new(File.join(TESTDIR, "numbers1.ods")) }
    assert_raises(TypeError) { Roo::Excelx.new(File.join(TESTDIR, "numbers1.xls")) }
  end

  def test_file_warning_error
    %w(ods xls).each do |extension|
      assert_raises(TypeError) do
        options = { packed: false, file_warning: :error }
        Roo::Excelx.new(File.join(TESTDIR, "numbers1. #{extension}"), options)
      end
    end
  end

  def test_file_warning_warning
    options = { packed: false, file_warning: :warning }
    assert_raises(ArgumentError) do
      Roo::Excelx.new(File.join(TESTDIR, "numbers1.ods"), options)
    end
  end

  def test_file_warning_ignore
    options = { packed: false, file_warning: :ignore }
    sheet = Roo::Excelx.new(File.join(TESTDIR, "type_excelx.ods"), options)
    assert sheet, "Should not throw an error"
  end

  def test_bug_xlsx_reference_cell
    # NOTE: If cell A contains a string and cell B references cell A.  When
    #       reading the value of cell B, the result will be "0.0" instead of the
    #       value of cell A.
    #
    # Before this test case, the following code:
    #
    # spreadsheet = Roo::Excelx.new("formula_string_error.xlsx")
    # spreadsheet.default_sheet = "sheet1"
    # p "A: #{spreadsheet.cell(1, 1)}" #=> "A: TestString"
    # p "B: #{spreadsheet.cell(2, 1)}" #=> "B: 0.0"
    #
    # where the expected result is
    # "A: TestString"
    # "B: TestString"
    xlsx = Roo::Excelx.new(File.join(TESTDIR, "formula_string_error.xlsx"))
    xlsx.default_sheet = xlsx.sheets.first
    assert_equal "Teststring", xlsx.cell("a", 1)
    assert_equal "Teststring", xlsx.cell("a", 2)
  end

  def test_parsing_xslx_from_numbers
    xlsx = Roo::Excelx.new(File.join(TESTDIR, "numbers-export.xlsx"))

    xlsx.default_sheet = xlsx.sheets.first
    assert_equal "Sheet 1", xlsx.cell("a", 1)

    # Another buggy behavior of Numbers 3.1: if a warkbook has more than a
    # single sheet, all sheets except the first one will have an extra row and
    # column added to the beginning. That's why we assert against cell B2 and
    # not A1
    xlsx.default_sheet = xlsx.sheets.last
    assert_equal "Sheet 2", xlsx.cell("b", 2)
  end

  def assert_cell_range_values(sheet, row_range, column_range, is_merged_range, expected_value)
    row_range.each do |row|
      column_range.each do |col|
        value = sheet.cell(col, row)
        if is_merged_range.call(row, col)
          assert_equal expected_value, value
        else
          assert_nil value
        end
      end
    end
  end

  def test_expand_merged_range
    options = { expand_merged_ranges: true }
    xlsx = Roo::Excelx.new(File.join(TESTDIR, "merged_ranges.xlsx"), options)

    [
      {
        rows: (3..7),
        columns: ("a".."b"),
        conditional: ->(row, col) { row > 3 && row < 7 && col == "a" },
        expected_value: "vertical1"
      },
      {
        rows: (3..11),
        columns: ("f".."h"),
        conditional: ->(row, col) { row > 3 && row < 11 && col == "g" },
        expected_value: "vertical2"
      },
      {
        rows: (3..5),
        columns: ("b".."f"),
        conditional: ->(row, col) { row == 4 && col > "b" && col < "f" },
        expected_value: "horizontal"
      },
      {
        rows: (8..13),
        columns: ("a".."e"),
        conditional: ->(row, col) { row > 8 && row < 13 && col > "a" && col < "e" },
        expected_value: "block"
      }
    ].each do |data|
      rows, cols, conditional, expected_value = data.values
      assert_cell_range_values(xlsx, rows, cols, conditional, expected_value)
    end
  end

  def test_noexpand_merged_range
    xlsx = Roo::Excelx.new(File.join(TESTDIR, "merged_ranges.xlsx"))

    [
      {
        rows: (3..7),
        columns: ("a".."b"),
        conditional: ->(row, col) { row == 4 && col == "a" },
        expected_value: "vertical1"
      },
      {
        rows: (3..11),
        columns: ("f".."h"),
        conditional: ->(row, col) { row == 4 && col == "g" },
        expected_value: "vertical2"
      },
      {
        rows: (3..5),
        columns: ("b".."f"),
        conditional: ->(row, col) { row == 4 && col == "c" },
        expected_value: "horizontal"
      },
      {
        rows: (8..13),
        columns: ("a".."e"),
        conditional: ->(row, col) { row == 9 && col == "b" },
        expected_value: "block"
      }
    ].each do |data|
      rows, cols, conditional, expected_value = data.values
      assert_cell_range_values(xlsx, rows, cols, conditional, expected_value)
    end
  end

  def test_open_stream
    file = filename(:numbers1)
    file_contents = File.read File.join(TESTDIR, file), encoding: "BINARY"
    stream = StringIO.new(file_contents)
    xlsx = Roo::Excelx.new(stream)
    expected_sheet_names = ["Tabelle1", "Name of Sheet 2", "Sheet3", "Sheet4", "Sheet5"]
    assert_equal expected_sheet_names, xlsx.sheets
  end

  def roo_class
    Roo::Excelx
  end

  def filename(name)
    "#{name}.xlsx"
  end
end
