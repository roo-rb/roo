require "test_helper"

class TestRooCSV < Minitest::Test
  def test_sheets
    file = filename("numbers1")
    workbook = roo_class.new(File.join(TESTDIR, file))
    assert_equal ["default"], workbook.sheets
    assert_raises(RangeError) { workbook.default_sheet = "no_sheet" }
    assert_raises(TypeError)  { workbook.default_sheet = [1, 2, 3] }
    workbook.sheets.each do |sh|
      workbook.default_sheet = sh
      assert_equal sh, workbook.default_sheet
    end
  end

  def test_download_uri_with_query_string
    file = filename("simple_spreadsheet")
    port = 12_347
    url = "#{local_server(port)}/#{file}?query-param=value"

    start_local_server(file, port) do
      csv = roo_class.new(url)
      assert_equal "Task 1", csv.cell("f", 4)
      assert_equal 1, csv.first_row
      assert_equal 13, csv.last_row
      assert_equal 1, csv.first_column
      assert_equal 6, csv.last_column
    end
  end

  def test_open_stream
    file = filename("Bibelbund")
    file_contents = File.read File.join(TESTDIR, file)
    stream = StringIO.new(file_contents)
    csv = roo_class.new(stream)

    assert_equal "Aktuelle Seite", csv.cell("h", 12)
    assert_equal 1, csv.first_row
    assert_equal 3735, csv.last_row
    assert_equal 1, csv.first_column
    assert_equal 8, csv.last_column
  end

  def test_nil_rows_and_lines_csv
    # x_123
    oo = Roo::CSV.new(File.join(TESTDIR,'Bibelbund.csv'))
    oo.default_sheet = oo.sheets.first
    assert_equal 1, oo.first_row
    assert_equal 3735, oo.last_row
    assert_equal 1, oo.first_column
    assert_equal 8, oo.last_column
  end

  def test_empty_csv
    # x_123
    oo = Roo::CSV.new(File.join(TESTDIR,'emptysheets.csv'))
    oo.default_sheet = oo.sheets.first
    assert_equal 1, oo.first_row
    assert_equal 1, oo.last_row
    assert_equal 1, oo.first_column
    assert_equal 1, oo.last_column
  end

  def test_csv_parsing_with_headers
    return unless CSV
    headers = ["TITEL", "VERFASSER", "OBJEKT", "NUMMER", "SEITE", "INTERNET", "PC", "KENNUNG"]

    oo = Roo::Spreadsheet.open(File.join(TESTDIR, "Bibelbund.csv"))
    parsed = oo.parse(headers: true)
    assert_equal headers, parsed[1].keys
  end

  def test_iso_8859_1
    file = File.open(File.join(TESTDIR, "iso_8859_1.csv"))
    options = { csv_options: { col_sep: ";", encoding: Encoding::ISO_8859_1 } }
    workbook = Roo::CSV.new(file.path, options)
    result = workbook.last_column
    assert_equal(19, result)
  end

  def roo_class
    Roo::CSV
  end

  def filename(name)
    "#{name}.csv"
  end
end
