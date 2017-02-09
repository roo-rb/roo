module TestLabels
  def test_labels
    options = { name: "named_cells", format: [:openoffice, :excelx, :libreoffice] }
    expected_labels = [
      ["anton", [5, 3, "Sheet1"]],
      ["berta", [4, 2, "Sheet1"]],
      ["caesar", [7, 2, "Sheet1"]],
    ]
    with_each_spreadsheet(options) do |oo|
      assert_equal expected_labels, oo.labels, "error with labels array in class #{oo.class}"
    end
  end

  def test_labeled_cells
    options = { name: "named_cells", format: [:openoffice, :excelx, :libreoffice] }
    with_each_spreadsheet(options) do |oo|
      oo.default_sheet = oo.sheets.first
      begin
        row, col = oo.label("anton")
      rescue ArgumentError
        puts "labels error at #{oo.class}"
        raise
      end
      assert_equal 5, row
      assert_equal 3, col

      row, col = oo.label("anton")
      assert_equal "Anton", oo.cell(row, col)

      row, col = oo.label("berta")
      assert_equal "Bertha", oo.cell(row, col)

      row, col = oo.label("caesar")
      assert_equal "Cäsar", oo.cell(row, col)

      row, col = oo.label("never")
      assert_nil row
      assert_nil col

      row, col, sheet = oo.label("anton")
      assert_equal 5, row
      assert_equal 3, col
      assert_equal "Sheet1", sheet

      assert_equal "Anton", oo.anton
      assert_raises(NoMethodError) do
        row, col = oo.never
      end

      # Reihenfolge row, col,sheet analog zu #label
      expected_labels = [
        ["anton", [5, 3, "Sheet1"]],
        ["berta", [4, 2, "Sheet1"]],
        ["caesar", [7, 2, "Sheet1"]],
      ]
      assert_equal expected_labels, oo.labels, "error with labels array in class #{oo.class}"
    end
  end

  def test_label
    options = { name: "named_cells", format: [:openoffice, :excelx, :libreoffice] }
    with_each_spreadsheet(options) do |oo|
      begin
        row, col = oo.label("anton")
      rescue ArgumentError
        puts "labels error at #{oo.class}"
        raise
      end

      assert_equal 5, row, "error with label in class #{oo.class}"
      assert_equal 3, col, "error with label in class #{oo.class}"

      row, col = oo.label("anton")
      assert_equal "Anton", oo.cell(row, col), "error with label in class #{oo.class}"

      row, col = oo.label("berta")
      assert_equal "Bertha", oo.cell(row, col), "error with label in class #{oo.class}"

      row, col = oo.label("caesar")
      assert_equal "Cäsar", oo.cell(row, col), "error with label in class #{oo.class}"

      row, col = oo.label("never")
      assert_nil row
      assert_nil col

      row, col, sheet = oo.label("anton")
      assert_equal 5, row
      assert_equal 3, col
      assert_equal "Sheet1", sheet
    end
  end

  def test_method_missing_anton
    options = { name: "named_cells", format: [:openoffice, :excelx, :libreoffice] }
    with_each_spreadsheet(options) do |oo|
      # oo.default_sheet = oo.sheets.first
      assert_equal "Anton", oo.anton
      assert_raises(NoMethodError) do
        oo.never
      end
    end
  end
end
