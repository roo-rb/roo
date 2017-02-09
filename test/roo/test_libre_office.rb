require "test_helper"

class TestRooOpenOffice < Minitest::Test
  def test_libre_office
    oo = Roo::LibreOffice.new(File.join(TESTDIR, "numbers1.ods"))
    oo.default_sheet = oo.sheets.first
    assert_equal 41, oo.cell("a", 12)
  end
end
