require "test_helper"

class TestRooFormatterCSV < Minitest::Test
  def test_date_time_to_csv
    with_each_spreadsheet(name: "time-test") do |workbook|
      Dir.mktmpdir do |tempdir|
        csv_output = File.join(tempdir, "time_test.csv")
        assert workbook.to_csv(csv_output)
        assert File.exist?(csv_output)
        assert_equal "", `diff --strip-trailing-cr #{TESTDIR}/time-test.csv #{csv_output}`
        # --strip-trailing-cr is needed because the test-file use 0A and
        # the test on an windows box generates 0D 0A as line endings
      end
    end
  end

  def test_boolean_to_csv
    with_each_spreadsheet(name: "boolean") do |workbook|
      Dir.mktmpdir do |tempdir|
        csv_output = File.join(tempdir,"boolean.csv")
        assert workbook.to_csv(csv_output)
        assert File.exist?(csv_output)
        assert_equal "", `diff --strip-trailing-cr #{TESTDIR}/boolean.csv #{csv_output}`
        # --strip-trailing-cr is needed because the test-file use 0A and
        # the test on an windows box generates 0D 0A as line endings
      end
    end
  end

  def test_link_to_csv
    with_each_spreadsheet(name: "link", format: :excelx) do |workbook|
      Dir.mktmpdir do |tempdir|
        csv_output = File.join(tempdir, "link.csv")
        assert workbook.to_csv(csv_output)
        assert File.exist?(csv_output)
        assert_equal "", `diff --strip-trailing-cr #{TESTDIR}/link.csv #{csv_output}`
        # --strip-trailing-cr is needed because the test-file use 0A and
        # the test on an windows box generates 0D 0A as line endings
      end
    end
  end

  # "/tmp/xxxx" darf man unter Windows nicht verwenden, weil das nicht erkannt
  # wird.
  # Besser: Methode um temporaeres Dir. portabel zu bestimmen
  def test_huge_document_to_csv
    skip_long_test

    original_csv_path = File.join(TESTDIR, "Bibelbund.csv")
    with_each_spreadsheet(name: "Bibelbund", format: [:openoffice, :excelx]) do |workbook|
      Dir.mktmpdir do |tempdir|
        new_csv_path = File.join(tempdir, "Bibelbund.csv")
        assert_equal "Tagebuch des Sekret\303\244rs.    Letzte Tagung 15./16.11.75 Schweiz", workbook.cell(45, "A")
        assert_equal "Tagebuch des Sekret\303\244rs.  Nachrichten aus Chile", workbook.cell(46, "A")
        assert_equal "Tagebuch aus Chile  Juli 1977", workbook.cell(55, "A")
        assert workbook.to_csv(new_csv_path)
        assert File.exist?(new_csv_path)
        assert FileUtils.identical?(original_csv_path, new_csv_path), "error in class #{workbook.class}"
      end
    end
  end

  def test_bug_empty_sheet
    with_each_spreadsheet(name: "formula", format: [:openoffice, :excelx]) do |workbook|
      workbook.default_sheet = "Sheet3" # is an empty sheet
      Dir.mktmpdir do |tempdir|
        workbook.to_csv(File.join(tempdir, "emptysheet.csv"))
        assert_equal "", `cat #{File.join(tempdir, "emptysheet.csv")}`
      end
    end
  end

  def test_bug_quotes_excelx
    skip_long_test
    # TODO: run this test with a much smaller document
    with_each_spreadsheet(name: "Bibelbund", format: [:openoffice, :excelx]) do |workbook|
      workbook.default_sheet = workbook.sheets.first
      assert_equal(
        'Einflüsse der neuen Theologie in "de gereformeerde Kerken van Nederland"',
        workbook.cell("A", 76)
      )
      workbook.to_csv("csv#{$$}")
      assert_equal(
        'Einflüsse der neuen Theologie in "de gereformeerde Kerken van Nederland"',
        workbook.cell("A", 78)
      )
      File.delete_if_exist("csv#{$$}")
    end
  end

  def test_bug_datetime_to_csv
    with_each_spreadsheet(name: "datetime") do |workbook|
      Dir.mktmpdir do |tempdir|
        datetime_csv_file = File.join(tempdir, "datetime.csv")

        assert workbook.to_csv(datetime_csv_file)
        assert File.exist?(datetime_csv_file)
        assert_equal "", file_diff("#{TESTDIR}/so_datetime.csv", datetime_csv_file)
      end
    end
  end

  def test_bug_datetime_offset_change
    # DO NOT REMOVE Asia/Calcutta
    [nil, "US/Eastern", "US/Pacific", "Asia/Calcutta"].each do |zone|
      with_timezone(zone) do
        with_each_spreadsheet(name: "datetime_timezone_ist_offset_change", format: %i[excelx openoffice libreoffice]) do |workbook|
          Dir.mktmpdir do |tempdir|
            datetime_csv_file = File.join(tempdir, "datetime_timezone_ist_offset_change.csv")

            assert workbook.to_csv(datetime_csv_file)
            assert File.exist?(datetime_csv_file)
            assert_equal "", file_diff("#{TESTDIR}/so_datetime_timezone_ist_offset_change.csv", datetime_csv_file)
          end
        end
      end
    end
  end

  def test_true_class
    assert_equal "true", cell_to_csv(1, 1)
  end

  def test_false_class
    assert_equal "false", cell_to_csv(2, 1)
  end

  def test_date_class
    assert_equal "2017-01-01", cell_to_csv(3, 1)
  end

  def cell_to_csv(row, col)
    filename = File.join(TESTDIR, "formula_cell_types.xlsx")
    Roo::Spreadsheet.open(filename).send("cell_to_csv", row, col, "Sheet1")
  end
end
