require "test_helper"

class TestRooFormatterYAML < Minitest::Test
  def test_date_time_yaml
    name = "time-test"
    expected = File.open(TESTDIR + "/expected_results/#{name}.yml").read
    with_each_spreadsheet(name: name) do |workbook|
      assert_equal expected, workbook.to_yaml
    end
  end

  def test_bug_to_yaml_empty_sheet
    formats = [:openoffice, :excelx]
    with_each_spreadsheet(name: "emptysheets", format: formats) do |workbook|
      workbook.default_sheet = workbook.sheets.first
      workbook.to_yaml
      assert_equal "", workbook.to_yaml
    end
  end
end
