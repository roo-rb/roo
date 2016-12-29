require "test_helper"
class TestRooOpenOffice < Minitest::Test
  def test_openoffice_download_uri_and_zipped
    start_local_server('rata.ods.zip') do
      url = 'http://0.0.0.0:5000/rata.ods.zip'
      oo = Roo::OpenOffice.new(url, packed: :zip)
      #has been changed: assert_equal 'ist "e" im Nenner von H(s)', sheet.cell('b', 5)
      assert_in_delta 0.001, 505.14, oo.cell('c', 33).to_f
    end
  end

  def test_download_uri_with_invalid_host
    assert_raises(RuntimeError) {
      Roo::OpenOffice.new("http://example.com/file.ods")
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
    Roo::OpenOffice
  end

  def filename(name)
    "#{name}.ods"
  end
end
