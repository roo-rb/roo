module TestFormulas
  def test_empty_sheet_formulas
    options = { name: "emptysheets", format: [:openoffice, :excelx] }
    with_each_spreadsheet(options) do |oo|
      oo.default_sheet = oo.sheets.first
      assert_equal [], oo.formulas, "An empty sheet's formulas should be an empty array"
    end
  end
end
