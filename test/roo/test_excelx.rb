require 'test_helper'

class TestRooExcelx < Minitest::Test
  def test_download_uri_with_invalid_host
    assert_raises(RuntimeError) {
      Roo::Excelx.new("http://example.com/file.xlsx")
    }
  end

  def test_download_uri_with_query_string
    file = filename("simple_spreadsheet")
    url = "#{TEST_URL}/#{file}?query-param=value"

    start_local_server(file) do
      spreadsheet = roo_class.new(url)
      spreadsheet.default_sheet = spreadsheet.sheets.first
      assert_equal 'Task 1', spreadsheet.cell('f', 4)
    end
  end

  def roo_class
    Roo::Excelx
  end

  def filename(name)
    "#{name}.xlsx"
  end
end
