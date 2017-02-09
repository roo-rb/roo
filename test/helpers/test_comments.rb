module TestComments
  def test_comment
    options = { name: "comments", format: [:openoffice, :libreoffice, :excelx] }
    with_each_spreadsheet(options) do |oo|
      assert_equal "Kommentar fuer B4", oo.comment("b", 4)
      assert_equal "Kommentar fuer B5", oo.comment("b", 5)
    end
  end

  def test_no_comment
    options = { name: "comments", format: [:openoffice, :excelx] }
    with_each_spreadsheet(options) do |oo|
      # There are no comments at the second sheet.
      assert_nil oo.comment("b", 4, oo.sheets[1])
    end
  end

  def test_comments
    options = { name: "comments", format: [:openoffice, :libreoffice, :excelx] }
    expexted_comments = [
      [4, 2, "Kommentar fuer B4"],
      [5, 2, "Kommentar fuer B5"],
    ]

    with_each_spreadsheet(options) do |oo|
      assert_equal expexted_comments, oo.comments(oo.sheets.first), "comments error in class #{oo.class}"
    end
  end

  def test_empty_sheet_comments
    options = { name: "emptysheets", format: [:openoffice, :excelx] }
    with_each_spreadsheet(options) do |oo|
      assert_equal [], oo.comments, "An empty sheet's formulas should be an empty array"
    end
  end

  def test_comments_from_google_sheet_exported_as_xlsx
    expected_comments = [[1, 1, "this is a comment\n\t-Steven Daniels"]]
    with_each_spreadsheet(name: "comments-google", format: [:excelx]) do |oo|
      assert_equal expected_comments, oo.comments(oo.sheets.first), "comments error in class #{oo.class}"
    end
  end
end
