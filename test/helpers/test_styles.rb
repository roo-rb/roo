module TestStyles
  def test_cell_styles
    # styles only valid in excel spreadsheets?
    # TODO: what todo with other spreadsheet types
    with_each_spreadsheet(name: "style", format: [:excelx]) do |oo|
      # bold
      assert_equal true, oo.font(1, 1).bold?
      assert_equal false, oo.font(1, 1).italic?
      assert_equal false, oo.font(1, 1).underline?

      # italic
      assert_equal false, oo.font(2, 1).bold?
      assert_equal true, oo.font(2, 1).italic?
      assert_equal false, oo.font(2, 1).underline?

      # normal
      assert_equal false, oo.font(3, 1).bold?
      assert_equal false, oo.font(3, 1).italic?
      assert_equal false, oo.font(3, 1).underline?

      # underline
      assert_equal false, oo.font(4, 1).bold?
      assert_equal false, oo.font(4, 1).italic?
      assert_equal true, oo.font(4, 1).underline?

      # bold italic
      assert_equal true, oo.font(5, 1).bold?
      assert_equal true, oo.font(5, 1).italic?
      assert_equal false, oo.font(5, 1).underline?

      # bold underline
      assert_equal true, oo.font(6, 1).bold?
      assert_equal false, oo.font(6, 1).italic?
      assert_equal true, oo.font(6, 1).underline?

      # italic underline
      assert_equal false, oo.font(7, 1).bold?
      assert_equal true, oo.font(7, 1).italic?
      assert_equal true, oo.font(7, 1).underline?

      # bolded row
      assert_equal true, oo.font(8, 1).bold?
      assert_equal false, oo.font(8, 1).italic?
      assert_equal false, oo.font(8, 1).underline?

      # bolded col
      assert_equal true, oo.font(9, 2).bold?
      assert_equal false, oo.font(9, 2).italic?
      assert_equal false, oo.font(9, 2).underline?

      # bolded row, italic col
      assert_equal true, oo.font(10, 3).bold?
      assert_equal true, oo.font(10, 3).italic?
      assert_equal false, oo.font(10, 3).underline?

      # normal
      assert_equal false, oo.font(11, 4).bold?
      assert_equal false, oo.font(11, 4).italic?
      assert_equal false, oo.font(11, 4).underline?
    end
  end
end
