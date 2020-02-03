# encoding: utf-8
require "test_helper"

class TestRooOpenOffice < Minitest::Test
  def test_openoffice_download_uri_and_zipped
    port = 12_345
    file = "rata.ods.zip"
    start_local_server(file, port) do
      url = "#{local_server(port)}/#{file}"

      workbook = roo_class.new(url, packed: :zip)
      assert_in_delta 0.001, 505.14, workbook.cell("c", 33).to_f
    end
  end

  def test_download_uri_with_invalid_host
    assert_raises(RuntimeError) do
      roo_class.new("http://examples.com/file.ods")
    end
  end

  def test_download_uri_with_query_string
    file = filename("simple_spreadsheet")
    port = 12_346
    url = "#{local_server(port)}/#{file}?query-param=value"
    start_local_server(file, port) do
      spreadsheet = roo_class.new(url)
      assert_equal "Task 1", spreadsheet.cell("f", 4)
    end
  end

  def test_openoffice_zipped
    workbook = roo_class.new(File.join(TESTDIR, "bode-v1.ods.zip"), packed: :zip)
    assert workbook
    assert_equal 'ist "e" im Nenner von H(s)', workbook.cell("b", 5)
  end

  def test_should_raise_file_not_found_error
    assert_raises(IOError) do
      roo_class.new(File.join("testnichtvorhanden", "Bibelbund.ods"))
    end
  end

  def test_file_warning_default_is_error
    expected_message = "test/files/numbers1.xls is not an openoffice spreadsheet"
    assert_raises(TypeError, expected_message) do
      roo_class.new(File.join(TESTDIR, "numbers1.xls"))
    end

    assert_raises(TypeError) do
      roo_class.new(File.join(TESTDIR, "numbers1.xlsx"))
    end
  end

  def test_file_warning_error
    options = { packed: false, file_warning: :error }

    assert_raises(TypeError) do
      roo_class.new(File.join(TESTDIR, "numbers1.xls"), options)
    end

    assert_raises(TypeError) do
      roo_class.new(File.join(TESTDIR, "numbers1.xlsx"), options)
    end
  end

  def test_file_warning_warning
    assert_raises(ArgumentError) do
      options = { packed: false, file_warning: :warning }
      roo_class.new(File.join(TESTDIR, "numbers1.xlsx"), options)
    end
  end

  def test_file_warning_ignore
    options = { packed: false, file_warning: :ignore }
    assert roo_class.new(File.join(TESTDIR, "type_openoffice.xlsx"), options), "Should not throw an error"
  end

  def test_encrypted_file
    workbook = roo_class.new(File.join(TESTDIR, "encrypted-letmein.ods"), password: "letmein")
    assert_equal "Hello World", workbook.cell("a", 1)
  end

  def test_encrypted_file_requires_password
    assert_raises(ArgumentError) do
      roo_class.new(File.join(TESTDIR, "encrypted-letmein.ods"))
    end
  end

  def test_encrypted_file_with_incorrect_password
    assert_raises(ArgumentError) do
      roo_class.new(File.join(TESTDIR, "encrypted-letmein.ods"), password: "badpassword")
    end
  end

  # 2011-08-11
  def test_bug_openoffice_formula_missing_letters
    # NOTE: This document was created using LibreOffice. The formulas seem
    # different from a document created using OpenOffice.
    #
    # TODO: translate
    # Bei den OpenOffice-Dateien ist in diesem Feld in der XML-
    # Datei of: als Prefix enthalten, waehrend in dieser Datei
    # irgendetwas mit oooc: als Prefix verwendet wird.
    workbook = roo_class.new(File.join(TESTDIR, "dreimalvier.ods"))
    assert_equal "=SUM([.A1:.D1])", workbook.formula("e", 1)
    assert_equal "=SUM([.A2:.D2])", workbook.formula("e", 2)
    assert_equal "=SUM([.A3:.D3])", workbook.formula("e", 3)
    expected_formulas = [
      [1, 5, "=SUM([.A1:.D1])"],
      [2, 5, "=SUM([.A2:.D2])"],
      [3, 5, "=SUM([.A3:.D3])"],
    ]
    assert_equal expected_formulas, workbook.formulas
  end

  def test_header_with_brackets_open_office
    options = { name: "advanced_header", format: :openoffice }
    with_each_spreadsheet(options) do |workbook|
      parsed_head = workbook.parse(headers: true)
      assert_equal "Date(yyyy-mm-dd)", workbook.cell("A", 1)
      assert_equal parsed_head[0].keys, ["Date(yyyy-mm-dd)"]
      assert_equal parsed_head[0].values, ["Date(yyyy-mm-dd)"]
    end
  end

  def test_office_version
    with_each_spreadsheet(name: "numbers1", format: :openoffice) do |workbook|
      assert_equal "1.0", workbook.officeversion
    end
  end

  def test_bug_contiguous_cells
    with_each_spreadsheet(name: "numbers1", format: :openoffice) do |workbook|
      workbook.default_sheet = "Sheet4"
      assert_equal Date.new(2007, 06, 16), workbook.cell("a", 1)
      assert_equal 10, workbook.cell("b", 1)
      assert_equal 10, workbook.cell("c", 1)
      assert_equal 10, workbook.cell("d", 1)
      assert_equal 10, workbook.cell("e", 1)
    end
  end

  def test_italo_table
    with_each_spreadsheet(name: "simple_spreadsheet_from_italo", format: :openoffice) do |workbook|
      assert_equal  "1", workbook.cell("A", 1)
      assert_equal  "1", workbook.cell("B", 1)
      assert_equal  "1", workbook.cell("C", 1)
      assert_equal  1, workbook.cell("A", 2).to_i
      assert_equal  2, workbook.cell("B", 2).to_i
      assert_equal  1, workbook.cell("C", 2).to_i
      assert_equal  1, workbook.cell("A", 3)
      assert_equal  3, workbook.cell("B", 3)
      assert_equal  1, workbook.cell("C", 3)
      assert_equal  "A", workbook.cell("A", 4)
      assert_equal  "A", workbook.cell("B", 4)
      assert_equal  "A", workbook.cell("C", 4)
      assert_equal  0.01, workbook.cell("A", 5)
      assert_equal  0.01, workbook.cell("B", 5)
      assert_equal  0.01, workbook.cell("C", 5)
      assert_equal 0.03, workbook.cell("a", 5) + workbook.cell("b", 5) + workbook.cell("c", 5)

      # Cells values in row 1:
      assert_equal "1:string", [workbook.cell(1, 1), workbook.celltype(1, 1)].join(":")
      assert_equal "1:string", [workbook.cell(1, 2), workbook.celltype(1, 2)].join(":")
      assert_equal "1:string", [workbook.cell(1, 3), workbook.celltype(1, 3)].join(":")

      # Cells values in row 2:
      assert_equal "1:string", [workbook.cell(2, 1), workbook.celltype(2, 1)].join(":")
      assert_equal "2:string", [workbook.cell(2, 2), workbook.celltype(2, 2)].join(":")
      assert_equal "1:string", [workbook.cell(2, 3), workbook.celltype(2, 3)].join(":")

      # Cells values in row 3:
      assert_equal "1:float", [workbook.cell(3, 1), workbook.celltype(3, 1)].join(":")
      assert_equal "3:float", [workbook.cell(3, 2), workbook.celltype(3, 2)].join(":")
      assert_equal "1:float", [workbook.cell(3, 3), workbook.celltype(3, 3)].join(":")

      # Cells values in row 4:
      assert_equal "A:string", [workbook.cell(4, 1), workbook.celltype(4, 1)].join(":")
      assert_equal "A:string", [workbook.cell(4, 2), workbook.celltype(4, 2)].join(":")
      assert_equal "A:string", [workbook.cell(4, 3), workbook.celltype(4, 3)].join(":")

      # Cells values in row 5:
      assert_equal "0.01:percentage", [workbook.cell(5, 1), workbook.celltype(5, 1)].join(":")
      assert_equal "0.01:percentage", [workbook.cell(5, 2), workbook.celltype(5, 2)].join(":")
      assert_equal "0.01:percentage", [workbook.cell(5, 3), workbook.celltype(5, 3)].join(":")
    end
  end

  def test_formula_openoffice
    with_each_spreadsheet(name: "formula", format: :openoffice) do |workbook|
      assert_equal 1, workbook.cell("A", 1)
      assert_equal 2, workbook.cell("A", 2)
      assert_equal 3, workbook.cell("A", 3)
      assert_equal 4, workbook.cell("A", 4)
      assert_equal 5, workbook.cell("A", 5)
      assert_equal 6, workbook.cell("A", 6)
      assert_equal 21, workbook.cell("A", 7)
      assert_equal :formula, workbook.celltype("A", 7)
      assert_equal "=[Sheet2.A1]", workbook.formula("C", 7)
      assert_nil workbook.formula("A", 6)
      expected_formulas = [
        [7, 1, "=SUM([.A1:.A6])"],
        [7, 2, "=SUM([.$A$1:.B6])"],
        [7, 3, "=[Sheet2.A1]"],
        [8, 2, "=SUM([.$A$1:.B7])"],
      ]
      assert_equal expected_formulas, workbook.formulas(workbook.sheets.first)

      # setting a cell
      workbook.set("A", 15, 41)
      assert_equal 41, workbook.cell("A", 15)
      workbook.set("A", 16, "41")
      assert_equal "41", workbook.cell("A", 16)
      workbook.set("A", 17, 42.5)
      assert_equal 42.5, workbook.cell("A", 17)
    end
  end

  def test_bug_ric
    with_each_spreadsheet(name: "ric", format: :openoffice) do |workbook|
      assert workbook.empty?("A", 1)
      assert workbook.empty?("B", 1)
      assert workbook.empty?("C", 1)
      assert workbook.empty?("D", 1)
      expected = 1
      letter = "e"
      while letter <= "u"
        assert_equal expected, workbook.cell(letter, 1)
        letter.succ!
        expected += 1
      end
      assert_equal "J", workbook.cell("v", 1)
      assert_equal "P", workbook.cell("w", 1)
      assert_equal "B", workbook.cell("x", 1)
      assert_equal "All", workbook.cell("y", 1)
      assert_equal 0, workbook.cell("a", 2)
      assert workbook.empty?("b", 2)
      assert workbook.empty?("c", 2)
      assert workbook.empty?("d", 2)
      assert_equal "B", workbook.cell("e", 2)
      assert_equal "B", workbook.cell("f", 2)
      assert_equal "B", workbook.cell("g", 2)
      assert_equal "B", workbook.cell("h", 2)
      assert_equal "B", workbook.cell("i", 2)
      assert_equal "B", workbook.cell("j", 2)
      assert_equal "B", workbook.cell("k", 2)
      assert_equal "B", workbook.cell("l", 2)
      assert_equal "B", workbook.cell("m", 2)
      assert_equal "B", workbook.cell("n", 2)
      assert_equal "B", workbook.cell("o", 2)
      assert_equal "B", workbook.cell("p", 2)
      assert_equal "B", workbook.cell("q", 2)
      assert_equal "B", workbook.cell("r", 2)
      assert_equal "B", workbook.cell("s", 2)
      assert workbook.empty?("t", 2)
      assert workbook.empty?("u", 2)
      assert_equal 0, workbook.cell("v", 2)
      assert_equal 0, workbook.cell("w", 2)
      assert_equal 15, workbook.cell("x", 2)
      assert_equal 15, workbook.cell("y", 2)
    end
  end

  def test_mehrteilig
    with_each_spreadsheet(name: "Bibelbund1", format: :openoffice) do |workbook|
      assert_equal "Tagebuch des Sekret\303\244rs.    Letzte Tagung 15./16.11.75 Schweiz", workbook.cell(45, "A")
    end
  end

  def test_cell_openoffice_html_escape
    with_each_spreadsheet(name: "html-escape", format: :openoffice) do |workbook|
      assert_equal "'", workbook.cell(1, 1)
      assert_equal "&", workbook.cell(2, 1)
      assert_equal ">", workbook.cell(3, 1)
      assert_equal "<", workbook.cell(4, 1)
      assert_equal "`", workbook.cell(5, 1)
      # test_openoffice_zipped will catch issues with &quot;
    end
  end

  def roo_class
    Roo::OpenOffice
  end

  def filename(name)
    "#{name}.ods"
  end
end
