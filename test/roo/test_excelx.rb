require "test_helper"

class TestRworkbookExcelx < Minitest::Test
  def test_download_uri_with_invalid_host
    assert_raises(RuntimeError) do
      Roo::Excelx.new("http://examples.com/file.xlsx")
    end
  end

  def test_download_uri_with_query_string
    file = filename("simple_spreadsheet")
    port = 12_344
    url = "#{local_server(port)}/#{file}?query-param=value"

    start_local_server(file, port) do
      spreadsheet = roo_class.new(url)
      assert_equal "Task 1", spreadsheet.cell("f", 4)
    end
  end

  def test_should_raise_file_not_found_error
    assert_raises(IOError) do
      roo_class.new(File.join("testnichtvorhanden", "Bibelbund.xlsx"))
    end
  end

  def test_file_warning_default
    assert_raises(TypeError) { roo_class.new(File.join(TESTDIR, "numbers1.ods")) }
    assert_raises(TypeError) { roo_class.new(File.join(TESTDIR, "numbers1.xls")) }
  end

  def test_file_warning_error
    %w(ods xls).each do |extension|
      assert_raises(TypeError) do
        options = { packed: false, file_warning: :error }
        roo_class.new(File.join(TESTDIR, "numbers1. #{extension}"), options)
      end
    end
  end

  def test_file_warning_warning
    options = { packed: false, file_warning: :warning }
    assert_raises(ArgumentError) do
      roo_class.new(File.join(TESTDIR, "numbers1.ods"), options)
    end
  end

  def test_file_warning_ignore
    options = { packed: false, file_warning: :ignore }
    sheet = roo_class.new(File.join(TESTDIR, "type_excelx.ods"), options)
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
    xlsx = roo_class.new(File.join(TESTDIR, "formula_string_error.xlsx"))
    assert_equal "Teststring", xlsx.cell("a", 1)
    assert_equal "Teststring", xlsx.cell("a", 2)
  end

  def test_parsing_xslx_from_numbers
    xlsx = roo_class.new(File.join(TESTDIR, "numbers-export.xlsx"))

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
    xlsx = roo_class.new(File.join(TESTDIR, "merged_ranges.xlsx"), options)

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
    xlsx = roo_class.new(File.join(TESTDIR, "merged_ranges.xlsx"))

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
    xlsx = roo_class.new(stream)
    expected_sheet_names = ["Tabelle1", "Name of Sheet 2", "Sheet3", "Sheet4", "Sheet5"]
    assert_equal expected_sheet_names, xlsx.sheets
  end

  def test_header_offset
    xlsx = roo_class.new(File.join(TESTDIR, "header_offset.xlsx"))
    data = xlsx.parse(column_1: "Header A1", column_2: "Header B1")
    assert_equal "Data A2", data[0][:column_1]
    assert_equal "Data B2", data[0][:column_2]
  end

  def test_formula_excelx
    with_each_spreadsheet(name: "formula", format: :excelx) do |workbook|
      assert_equal 1, workbook.cell("A", 1)
      assert_equal 2, workbook.cell("A", 2)
      assert_equal 3, workbook.cell("A", 3)
      assert_equal 4, workbook.cell("A", 4)
      assert_equal 5, workbook.cell("A", 5)
      assert_equal 6, workbook.cell("A", 6)
      assert_equal 21, workbook.cell("A", 7)
      assert_equal :formula, workbook.celltype("A", 7)
      assert_nil workbook.formula("A", 6)

      expected_result = [
        [7, 1, "SUM(A1:A6)"],
        [7, 2, "SUM($A$1:B6)"],
      ]
      assert_equal expected_result, workbook.formulas(workbook.sheets.first)

      # setting a cell
      workbook.set("A", 15, 41)
      assert_equal 41, workbook.cell("A", 15)
      workbook.set("A", 16, "41")
      assert_equal "41", workbook.cell("A", 16)
      workbook.set("A", 17, 42.5)
      assert_equal 42.5, workbook.cell("A", 17)
    end
  end

  # TODO: temporaerer Test
  def test_seiten_als_date
    skip_long_test

    with_each_spreadsheet(name: "Bibelbund", format: :excelx) do |workbook|
      assert_equal "Bericht aus dem Sekretariat", workbook.cell(13, 1)
      assert_equal "1981-4", workbook.cell(13, "D")
      assert_equal String, workbook.excelx_type(13, "E")[1].class
      assert_equal [:numeric_or_formula, "General"], workbook.excelx_type(13, "E")
      assert_equal "428", workbook.excelx_value(13, "E")
      assert_equal 428.0, workbook.cell(13, "E")
    end
  end

  def test_bug_simple_spreadsheet_time_bug
    # really a bug? are cells really of type time?
    # No! :float must be the correct type
    with_each_spreadsheet(name: "simple_spreadsheet", format: :excelx) do |workbook|
      # puts workbook.cell("B", 5).to_s
      # assert_equal :time, workbook.celltype("B", 5)
      assert_equal :float, workbook.celltype("B", 5)
      assert_equal 10.75, workbook.cell("B", 5)
      assert_equal 12.50, workbook.cell("C", 5)
      assert_equal 0, workbook.cell("D", 5)
      assert_equal 1.75, workbook.cell("E", 5)
      assert_equal "Task 1", workbook.cell("F", 5)
      assert_equal Date.new(2007, 5, 7), workbook.cell("A", 5)
    end
  end

  def test_simple2_excelx
    with_each_spreadsheet(name: "simple_spreadsheet", format: :excelx) do |workbook|
      assert_equal [:numeric_or_formula, "yyyy\\-mm\\-dd"], workbook.excelx_type("A", 4)
      assert_equal [:numeric_or_formula, "#,##0.00"], workbook.excelx_type("B", 4)
      assert_equal [:numeric_or_formula, "#,##0.00"], workbook.excelx_type("c", 4)
      assert_equal [:numeric_or_formula, "General"], workbook.excelx_type("d", 4)
      assert_equal [:numeric_or_formula, "General"], workbook.excelx_type("e", 4)
      assert_equal :string, workbook.excelx_type("f", 4)

      assert_equal "39209", workbook.excelx_value("a", 4)
      assert_equal "yyyy\\-mm\\-dd", workbook.excelx_format("a", 4)
      assert_equal "9.25", workbook.excelx_value("b", 4)
      assert_equal "10.25", workbook.excelx_value("c", 4)
      assert_equal "0", workbook.excelx_value("d", 4)
      # ... Sum-Spalte
      # assert_equal "Task 1", workbook.excelx_value("f", 4)
      assert_equal "Task 1", workbook.cell("f", 4)
      assert_equal Date.new(2007, 05, 07), workbook.cell("a", 4)
      assert_equal "9.25", workbook.excelx_value("b", 4)
      assert_equal "#,##0.00", workbook.excelx_format("b", 4)
      assert_equal 9.25, workbook.cell("b", 4)
      assert_equal :float, workbook.celltype("b", 4)
      assert_equal :float, workbook.celltype("d", 4)
      assert_equal 0, workbook.cell("d", 4)
      assert_equal :formula, workbook.celltype("e", 4)
      assert_equal 1, workbook.cell("e", 4)
      assert_equal "C4-B4-D4", workbook.formula("e", 4)
      assert_equal :string, workbook.celltype("f", 4)
      assert_equal "Task 1", workbook.cell("f", 4)
    end
  end

  def test_bug_pfand_from_windows_phone_xlsx
    # skip_jruby_incompatible_test
    # TODO: Does JRUBY need to skip this test
    return if defined? JRUBY_VERSION

    options = { name: "Pfand_from_windows_phone", format: :excelx }
    with_each_spreadsheet(options) do |workbook|
      workbook.default_sheet = workbook.sheets.first
      assert_equal ["Blatt1", "Blatt2", "Blatt3"], workbook.sheets
      assert_equal "Summe", workbook.cell("b", 1)

      assert_equal Date.new(2011, 9, 14), workbook.cell("a", 2)
      assert_equal :date, workbook.celltype("a", 2)
      assert_equal Date.new(2011, 9, 15), workbook.cell("a", 3)
      assert_equal :date, workbook.celltype("a", 3)

      assert_equal 3.81, workbook.cell("b", 2)
      assert_equal "SUM(C2:L2)", workbook.formula("b", 2)
      assert_equal 0.7, workbook.cell("c", 2)
    end # each
  end

  def test_excelx_links
    with_each_spreadsheet(name: "link", format: :excelx) do |workbook|
      assert_equal "Google", workbook.cell(1, 1)
      assert_equal "http://www.google.com", workbook.cell(1, 1).href
    end
  end

  def test_handles_link_without_hyperlink
    workbook = Roo::Spreadsheet.open(File.join(TESTDIR, "bad_link.xlsx"))
    assert_equal "Test", workbook.cell(1, 1)
  end

  # Excel has two base date formats one from 1900 and the other from 1904.
  # see #test_base_dates_in_excel
  def test_base_dates_in_excelx
    with_each_spreadsheet(name: "1900_base", format: :excelx) do |workbook|
      assert_equal Date.new(2009, 06, 15), workbook.cell(1, 1)
      assert_equal :date, workbook.celltype(1, 1)
    end
    with_each_spreadsheet(name: "1904_base", format: :excelx) do |workbook|
      assert_equal Date.new(2009, 06, 15), workbook.cell(1, 1)
      assert_equal :date, workbook.celltype(1, 1)
    end
  end

  def roo_class
    Roo::Excelx
  end

  def filename(name)
    "#{name}.xlsx"
  end
end
