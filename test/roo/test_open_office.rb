# encoding: utf-8
require "test_helper"

class TestRooOpenOffice < Minitest::Test
  def test_openoffice_download_uri_and_zipped
    port = 12_345
    file = "rata.ods.zip"
    start_local_server(file, port) do
      url = "#{local_server(port)}/#{file}"

      oo = roo_class.new(url, packed: :zip)
      assert_in_delta 0.001, 505.14, oo.cell("c", 33).to_f
    end
  end

  def test_download_uri_with_invalid_host
    assert_raises(RuntimeError) do
      roo_class.new("http://example.com/file.ods")
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
    oo = roo_class.new(File.join(TESTDIR, "bode-v1.ods.zip"), packed: :zip)
    assert oo
    assert_equal 'ist "e" im Nenner von H(s)', oo.cell("b", 5)
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
    assert_raises(TypeError) do
      roo_class.new(File.join(TESTDIR, "numbers1.xls"),
      packed: false,
      file_warning: :error
      )
    end

    assert_raises(TypeError) do
      roo_class.new(File.join(TESTDIR, "numbers1.xlsx"),
      packed: false,
      file_warning: :error)
    end
  end

  def test_file_warning_warning
    assert_raises(ArgumentError) do
      roo_class.new(File.join(TESTDIR, "numbers1.xlsx"),
      packed: false,
      file_warning: :warning)
    end
  end

  def test_file_warning_ignore
    assert roo_class.new(File.join(TESTDIR, "type_openoffice.xlsx"),
    packed: false,
    file_warning: :ignore), "Should not throw an error"
  end

  def test_encrypted_file
    oo = roo_class.new(File.join(TESTDIR, "encrypted-letmein.ods"), password: "letmein")
    assert_equal "Hello World", oo.cell("a", 1)
  end

  def test_encrypted_file_requires_password
    assert_raises(ArgumentError) { roo_class.new(File.join(TESTDIR, "encrypted-letmein.ods")) }
  end

  def test_encrypted_file_with_incorrect_password
    assert_raises(ArgumentError) { roo_class.new(File.join(TESTDIR, "encrypted-letmein.ods"), password:  "badpassword") }
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
    oo = Roo::OpenOffice.new(File.join(TESTDIR,'dreimalvier.ods'))
    assert_equal '=SUM([.A1:.D1])', oo.formula('e',1)
    assert_equal '=SUM([.A2:.D2])', oo.formula('e',2)
    assert_equal '=SUM([.A3:.D3])', oo.formula('e',3)
    assert_equal [
      [1,5,'=SUM([.A1:.D1])'],
      [2,5,'=SUM([.A2:.D2])'],
      [3,5,'=SUM([.A3:.D3])'],
      ], oo.formulas

  end

  def roo_class
    Roo::OpenOffice
  end

  def filename(name)
    "#{name}.ods"
  end
end
