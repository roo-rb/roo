# NOTE: Putting these tests into modules in order to share them across different
#       test classes, i.e. both TestRooExcelx and TestRooOpenOffice should run
#       sheet related tests.
#
#       This will allow me to reuse these test cases when I add new classes for
#       Roo's future API.
# Sheet related tests
module TestSheets
  def test_sheets
    sheet_names = ["Tabelle1", "Name of Sheet 2", "Sheet3", "Sheet4", "Sheet5"]
    with_each_spreadsheet(name: "numbers1") do |oo|
      assert_equal sheet_names, oo.sheets
      assert_raises(RangeError) { oo.default_sheet = "no_sheet" }
      assert_raises(TypeError)  { oo.default_sheet = [1, 2, 3] }
      oo.sheets.each do |sheet_name|
        oo.default_sheet = sheet_name
        assert_equal sheet_name, oo.default_sheet
      end
    end
  end

  def test_sheetname
    bad_sheet_name = "non existing sheet name"
    with_each_spreadsheet(name: "numbers1") do |oo|
      oo.default_sheet = "Name of Sheet 2"
      assert_equal "I am sheet 2", oo.cell("C", 5)
      assert_raises(RangeError) { oo.default_sheet = bad_sheet_name }
      assert_raises(RangeError) { oo.default_sheet = bad_sheet_name }
      assert_raises(RangeError) { oo.cell("C", 5, bad_sheet_name) }
      assert_raises(RangeError) { oo.celltype("C", 5, bad_sheet_name) }
      assert_raises(RangeError) { oo.empty?("C", 5, bad_sheet_name) }
      assert_raises(RangeError) { oo.formula?("C", 5, bad_sheet_name) }
      assert_raises(RangeError) { oo.formula("C", 5, bad_sheet_name) }
      assert_raises(RangeError) { oo.set("C", 5, 42, bad_sheet_name) }
      assert_raises(RangeError) { oo.formulas(bad_sheet_name) }
      assert_raises(RangeError) { oo.to_yaml({}, 1, 1, 1, 1, bad_sheet_name) }
    end
  end

  def test_info_doesnt_set_default_sheet
    sheet_name = "Sheet3"
    with_each_spreadsheet(name: "numbers1") do |oo|
      oo.default_sheet = sheet_name
      oo.info
      assert_equal sheet_name, oo.default_sheet
    end
  end

  def test_bug_numbered_sheet_names
    options = { name: "bug-numbered-sheet-names", format: :excelx }
    with_each_spreadsheet(options) do |oo|
      oo.each_with_pagename {}
    end
  end
end
