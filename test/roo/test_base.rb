require "test_helper"

class TestRooBase < Minitest::Test
  def test_info
    # NOTE: unfortunately, the ods and xlsx versions of numbers1 are not
    #       identical, so this test fails for Open Office.
    expected_templ = File.read("#{TESTDIR}/expected_results/numbers_info.yml")
    with_each_spreadsheet(name: "numbers1", format: [:excelx]) do |workbook|
      ext = get_extension(workbook)
      expected = Kernel.format(expected_templ, ext)
      assert_equal expected.strip, workbook.info.strip
    end
  end

  def test_column
    with_each_spreadsheet(name: "numbers1") do |workbook|
      expected = [1.0, 5.0, nil, 10.0, Date.new(1961, 11, 21), "tata", nil, nil, nil, nil, "thisisa11", 41.0, nil, nil, 41.0, "einundvierzig", nil, Date.new(2007, 5, 31)]
      assert_equal expected, workbook.column(1)
      assert_equal expected, workbook.column("a")
    end
  end

  def test_column_huge_document
    skip_long_test
    with_each_spreadsheet(name: "Bibelbund", format: [:openoffice, :excelx]) do |workbook|
      workbook.default_sheet = workbook.sheets.first
      assert_equal 3735, workbook.column("a").size
    end
  end

  def test_simple_spreadsheet_find_by_condition
    with_each_spreadsheet(name: "simple_spreadsheet") do |workbook|
      workbook.header_line = 3
      results = workbook.find(:all, conditions: { "Comment" => "Task 1" })
      assert_equal Date.new(2007, 05, 07), results[1]["Date"]
      assert_equal 10.75, results[1]["Start time"]
      assert_equal 12.50, results[1]["End time"]
      assert_equal 0, results[1]["Pause"]
      assert_equal 1.75, results[1]["Sum"]
      assert_equal "Task 1", results[1]["Comment"]
    end
  end

  def test_bug_bbu
    expected_templ = File.read("#{TESTDIR}/expected_results/bbu_info.txt")
    with_each_spreadsheet(name: "bbu", format: [:openoffice, :excelx]) do |workbook|
      ext = get_extension(workbook)
      expected_result = Kernel.format(expected_templ, ext)
      assert_equal expected_result.strip, workbook.info.strip

      workbook.default_sheet = workbook.sheets[1] # empty sheet
      assert_nil workbook.first_row
      assert_nil workbook.last_row
      assert_nil workbook.first_column
      assert_nil workbook.last_column
    end
  end

  def test_find_by_row_huge_document
    skip_long_test
    options = { name: "Bibelbund", format: [:openoffice, :excelx] }
    with_each_spreadsheet(options) do |workbook|
      workbook.default_sheet = workbook.sheets.first
      result = workbook.find 20
      assert result
      assert_equal "Brief aus dem Sekretariat", result[0]

      result = workbook.find 22
      assert result
      assert_equal "Brief aus dem Skretariat. Tagung in Amberg/Opf.", result[0]
    end
  end

  def test_find_by_row
    with_each_spreadsheet(name: "numbers1") do |workbook|
      workbook.header_line = nil
      result = workbook.find 16
      assert result
      assert_nil workbook.header_line
      # keine Headerlines in diesem Beispiel definiert
      assert_equal "einundvierzig", result[0]
      # assert_equal false, results
      result = workbook.find 15
      assert result
      assert_equal 41, result[0]
    end
  end

  def test_find_by_row_if_header_line_is_not_nil
    with_each_spreadsheet(name: "numbers1") do |workbook|
      workbook.header_line = 2
      refute_nil workbook.header_line
      results = workbook.find 1
      assert results
      assert_equal 5, results[0]
      assert_equal 6, results[1]
      results = workbook.find 15
      assert results
      assert_equal "einundvierzig", results[0]
    end
  end

  def test_find_by_conditions
    skip_long_test
    expected_results = [
      {
        "VERFASSER" => "Almassy, Annelene von",
        "INTERNET" => nil,
        "SEITE" => 316.0,
        "KENNUNG" => "Aus dem Bibelbund",
        "OBJEKT" => "Bibel+Gem",
        "PC" => "#C:\\Bibelbund\\reprint\\BuG1982-3.pdf#",
        "NUMMER" => "1982-3",
        "TITEL" => "Brief aus dem Sekretariat"
      },
      {
        "VERFASSER" => "Almassy, Annelene von",
        "INTERNET" => nil,
        "SEITE" => 222.0,
        "KENNUNG" => "Aus dem Bibelbund",
        "OBJEKT" => "Bibel+Gem",
        "PC" => "#C:\\Bibelbund\\reprint\\BuG1983-2.pdf#",
        "NUMMER" => "1983-2",
        "TITEL" => "Brief aus dem Sekretariat"
      }
    ]

    expected_results_size = 2
    options = { name: "Bibelbund", format: [:openoffice, :excelx] }
    with_each_spreadsheet(options) do |workbook|
      results = workbook.find(:all, conditions: { "TITEL" => "Brief aus dem Sekretariat" })
      assert_equal expected_results_size, results.size
      assert_equal expected_results, results

      conditions = {
        "TITEL" => "Brief aus dem Sekretariat",
        "VERFASSER" => "Almassy, Annelene von"
      }
      results = workbook.find(:all, conditions: conditions)
      assert_equal expected_results, results
      assert_equal expected_results_size, results.size

      results = workbook.find(:all, conditions: { "VERFASSER" => "Almassy, Annelene von" })
      assert_equal 13, results.size
    end
  end

  def test_find_by_conditions_with_array_option
    expected_results = [
      [
        "Brief aus dem Sekretariat",
        "Almassy, Annelene von",
        "Bibel+Gem",
        "1982-3",
        316.0,
        nil,
        "#C:\\Bibelbund\\reprint\\BuG1982-3.pdf#",
        "Aus dem Bibelbund",
      ],
      [
        "Brief aus dem Sekretariat",
        "Almassy, Annelene von",
        "Bibel+Gem",
        "1983-2",
        222.0,
        nil,
        "#C:\\Bibelbund\\reprint\\BuG1983-2.pdf#",
        "Aus dem Bibelbund",
      ]
    ]
    options = { name: "Bibelbund", format: [:openoffice, :excelx] }
    with_each_spreadsheet(options) do |workbook|
      conditions = {
        "TITEL" => "Brief aus dem Sekretariat",
        "VERFASSER" => "Almassy, Annelene von"
      }
      results = workbook.find(:all, conditions: conditions, array: true)
      assert_equal 2, results.size
      assert_equal expected_results, results
    end
  end
end
