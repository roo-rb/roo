
# Dump warnings that come from the test to open files
# with the wrong spreadsheet class
#STDERR.reopen "/dev/null","w"

Encoding.default_external = "UTF-8"

require 'test_helper'
require 'stringio'

class TestRoo < Minitest::Test
  FORMATS = [
    :excelx,
    :excelxm,
    :openoffice,
    :libreoffice
  ]

  LONG_RUN = false

  def fixture_filename(name, format)
    case format
    when :excelx
      "#{name}.xlsx"
    when :excelxm
      "#{name}.xlsm"
    when :openoffice, :libreoffice
      "#{name}.ods"
    else
      raise ArgumentError, "unexpected format #{format}"
    end
  end

  # call a block of code for each spreadsheet type
  # and yield a reference to the roo object
  def with_each_spreadsheet(options)
    if options[:format]
      formats = Array(options[:format])
      invalid_formats = formats - FORMATS
      unless invalid_formats.empty?
        raise "invalid spreadsheet types: #{invalid_formats.join(', ')}"
      end
    else
      formats = FORMATS
    end
    formats.each do |format|
      begin
        yield Roo::Spreadsheet.open(File.join(TESTDIR,
          fixture_filename(options[:name], format)))
      rescue => e
        raise e, "#{e.message} for #{format}", e.backtrace unless options[:ignore_errors]
      end
    end
  end

  def test_sheets
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      assert_equal ["Tabelle1","Name of Sheet 2","Sheet3","Sheet4","Sheet5"], oo.sheets
      assert_raises(RangeError) { oo.default_sheet = "no_sheet" }
      assert_raises(TypeError)  { oo.default_sheet = [1,2,3] }
      oo.sheets.each { |sh|
        oo.default_sheet = sh
        assert_equal sh, oo.default_sheet
      }
    end
  end

  def test_cells
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      # warum ist Auswaehlen erstes sheet hier nicht
      # mehr drin?
      oo.default_sheet = oo.sheets.first
      assert_equal 1, oo.cell(1,1)
      assert_equal 2, oo.cell(1,2)
      assert_equal 3, oo.cell(1,3)
      assert_equal 4, oo.cell(1,4)
      assert_equal 5, oo.cell(2,1)
      assert_equal 6, oo.cell(2,2)
      assert_equal 7, oo.cell(2,3)
      assert_equal 8, oo.cell(2,4)
      assert_equal 9, oo.cell(2,5)
      assert_equal "test", oo.cell(2,6)
      assert_equal :string, oo.celltype(2,6)
      assert_equal 11, oo.cell(2,7)
      unless oo.kind_of? Roo::CSV
        assert_equal :float, oo.celltype(2,7)
      end
      assert_equal 10, oo.cell(4,1)
      assert_equal 11, oo.cell(4,2)
      assert_equal 12, oo.cell(4,3)
      assert_equal 13, oo.cell(4,4)
      assert_equal 14, oo.cell(4,5)
      assert_equal 10, oo.cell(4,'A')
      assert_equal 11, oo.cell(4,'B')
      assert_equal 12, oo.cell(4,'C')
      assert_equal 13, oo.cell(4,'D')
      assert_equal 14, oo.cell(4,'E')
      unless oo.kind_of? Roo::CSV
        assert_equal :date, oo.celltype(5,1)
        assert_equal Date.new(1961,11,21), oo.cell(5,1)
        assert_equal "1961-11-21", oo.cell(5,1).to_s
      end
    end
  end

  def test_celltype
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      assert_equal :string, oo.celltype(2,6)
    end
  end

  def test_cell_address
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      assert_equal "tata", oo.cell(6,1)
      assert_equal "tata", oo.cell(6,'A')
      assert_equal "tata", oo.cell('A',6)
      assert_equal "tata", oo.cell(6,'a')
      assert_equal "tata", oo.cell('a',6)
      assert_raises(ArgumentError) { assert_equal "tata", oo.cell('a','f') }
      assert_raises(ArgumentError) { assert_equal "tata", oo.cell('f','a') }
      assert_equal "thisisc8", oo.cell(8,3)
      assert_equal "thisisc8", oo.cell(8,'C')
      assert_equal "thisisc8", oo.cell('C',8)
      assert_equal "thisisc8", oo.cell(8,'c')
      assert_equal "thisisc8", oo.cell('c',8)
      assert_equal "thisisd9", oo.cell('d',9)
      assert_equal "thisisa11", oo.cell('a',11)
    end
  end

  def test_office_version
    with_each_spreadsheet(:name=>'numbers1', :format=>:openoffice) do |oo|
      assert_equal "1.0", oo.officeversion
    end
  end

  def test_sheetname
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      oo.default_sheet = "Name of Sheet 2"
      assert_equal 'I am sheet 2', oo.cell('C',5)
      assert_raises(RangeError) { oo.default_sheet = "non existing sheet name" }
      assert_raises(RangeError) { oo.default_sheet = "non existing sheet name" }
      assert_raises(RangeError) { oo.cell('C',5,"non existing sheet name")}
      assert_raises(RangeError) { oo.celltype('C',5,"non existing sheet name")}
      assert_raises(RangeError) { oo.empty?('C',5,"non existing sheet name")}
      assert_raises(RangeError) { oo.formula?('C',5,"non existing sheet name")}
      assert_raises(RangeError) { oo.formula('C',5,"non existing sheet name")}
      assert_raises(RangeError) { oo.set('C',5,42,"non existing sheet name")}
      assert_raises(RangeError) { oo.formulas("non existing sheet name")}
      assert_raises(RangeError) { oo.to_yaml({},1,1,1,1,"non existing sheet name")}
    end
  end

  def test_argument_error
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      oo.default_sheet = "Tabelle1"
    end
  end

  def test_bug_contiguous_cells
    with_each_spreadsheet(:name=>'numbers1', :format=>:openoffice) do |oo|
      oo.default_sheet = "Sheet4"
      assert_equal Date.new(2007,06,16), oo.cell('a',1)
      assert_equal 10, oo.cell('b',1)
      assert_equal 10, oo.cell('c',1)
      assert_equal 10, oo.cell('d',1)
      assert_equal 10, oo.cell('e',1)
    end
  end

  def test_bug_italo_ve
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      oo.default_sheet = "Sheet5"
      assert_equal 1, oo.cell('A',1)
      assert_equal 5, oo.cell('b',1)
      assert_equal 5, oo.cell('c',1)
      assert_equal 2, oo.cell('a',2)
      assert_equal 3, oo.cell('a',3)
    end
  end

  def test_italo_table
    with_each_spreadsheet(:name=>'simple_spreadsheet_from_italo', :format=>:openoffice) do |oo|
      assert_equal  '1', oo.cell('A',1)
      assert_equal  '1', oo.cell('B',1)
      assert_equal  '1', oo.cell('C',1)
      assert_equal  1, oo.cell('A',2).to_i
      assert_equal  2, oo.cell('B',2).to_i
      assert_equal  1, oo.cell('C',2).to_i
      assert_equal  1, oo.cell('A',3)
      assert_equal  3, oo.cell('B',3)
      assert_equal  1, oo.cell('C',3)
      assert_equal  'A', oo.cell('A',4)
      assert_equal  'A', oo.cell('B',4)
      assert_equal  'A', oo.cell('C',4)
      assert_equal  0.01, oo.cell('A',5)
      assert_equal  0.01, oo.cell('B',5)
      assert_equal  0.01, oo.cell('C',5)
      assert_equal 0.03, oo.cell('a',5)+oo.cell('b',5)+oo.cell('c',5)

      # Cells values in row 1:
      assert_equal "1:string", oo.cell(1, 1)+":"+oo.celltype(1, 1).to_s
      assert_equal "1:string",oo.cell(1, 2)+":"+oo.celltype(1, 2).to_s
      assert_equal "1:string",oo.cell(1, 3)+":"+oo.celltype(1, 3).to_s

      # Cells values in row 2:
      assert_equal "1:string",oo.cell(2, 1)+":"+oo.celltype(2, 1).to_s
      assert_equal "2:string",oo.cell(2, 2)+":"+oo.celltype(2, 2).to_s
      assert_equal "1:string",oo.cell(2, 3)+":"+oo.celltype(2, 3).to_s

      # Cells values in row 3:
      assert_equal "1:float",oo.cell(3, 1).to_s+":"+oo.celltype(3, 1).to_s
      assert_equal "3:float",oo.cell(3, 2).to_s+":"+oo.celltype(3, 2).to_s
      assert_equal "1:float",oo.cell(3, 3).to_s+":"+oo.celltype(3, 3).to_s

      # Cells values in row 4:
      assert_equal "A:string",oo.cell(4, 1)+":"+oo.celltype(4, 1).to_s
      assert_equal "A:string",oo.cell(4, 2)+":"+oo.celltype(4, 2).to_s
      assert_equal "A:string",oo.cell(4, 3)+":"+oo.celltype(4, 3).to_s

      # Cells values in row 5:
      if oo.class == Roo::OpenOffice
        assert_equal "0.01:percentage",oo.cell(5, 1).to_s+":"+oo.celltype(5, 1).to_s
        assert_equal "0.01:percentage",oo.cell(5, 2).to_s+":"+oo.celltype(5, 2).to_s
        assert_equal "0.01:percentage",oo.cell(5, 3).to_s+":"+oo.celltype(5, 3).to_s
      else
        assert_equal "0.01:float",oo.cell(5, 1).to_s+":"+oo.celltype(5, 1).to_s
        assert_equal "0.01:float",oo.cell(5, 2).to_s+":"+oo.celltype(5, 2).to_s
        assert_equal "0.01:float",oo.cell(5, 3).to_s+":"+oo.celltype(5, 3).to_s
      end
    end
  end

  def test_formula_openoffice
    with_each_spreadsheet(:name=>'formula', :format=>:openoffice) do |oo|
      assert_equal 1, oo.cell('A',1)
      assert_equal 2, oo.cell('A',2)
      assert_equal 3, oo.cell('A',3)
      assert_equal 4, oo.cell('A',4)
      assert_equal 5, oo.cell('A',5)
      assert_equal 6, oo.cell('A',6)
      assert_equal 21, oo.cell('A',7)
      assert_equal :formula, oo.celltype('A',7)
      assert_equal "=[Sheet2.A1]", oo.formula('C',7)
      assert_nil oo.formula('A',6)
      assert_equal [[7, 1, "=SUM([.A1:.A6])"],
        [7, 2, "=SUM([.$A$1:.B6])"],
        [7, 3, "=[Sheet2.A1]"],
        [8, 2, "=SUM([.$A$1:.B7])"],
      ], oo.formulas(oo.sheets.first)

      # setting a cell
      oo.set('A',15, 41)
      assert_equal 41, oo.cell('A',15)
      oo.set('A',16, "41")
      assert_equal "41", oo.cell('A',16)
      oo.set('A',17, 42.5)
      assert_equal 42.5, oo.cell('A',17)
    end
  end

  def test_header_with_brackets_excelx
    with_each_spreadsheet(:name => 'advanced_header', :format => :openoffice) do |oo|
      parsed_head = oo.parse(:headers => true)
      assert_equal "Date(yyyy-mm-dd)", oo.cell('A',1)
      assert_equal parsed_head[0].keys, ["Date(yyyy-mm-dd)"]
      assert_equal parsed_head[0].values, ["Date(yyyy-mm-dd)"]
    end
  end

  def test_formula_excelx
    with_each_spreadsheet(:name=>'formula', :format=>:excelx) do |oo|
      assert_equal 1, oo.cell('A',1)
      assert_equal 2, oo.cell('A',2)
      assert_equal 3, oo.cell('A',3)
      assert_equal 4, oo.cell('A',4)
      assert_equal 5, oo.cell('A',5)
      assert_equal 6, oo.cell('A',6)
      assert_equal 21, oo.cell('A',7)
      assert_equal :formula, oo.celltype('A',7)
      #steht nicht in Datei, oder?
      #nein, diesen Bezug habe ich nur in der OpenOffice-Datei
      #assert_equal "=[Sheet2.A1]", oo.formula('C',7)
      assert_nil oo.formula('A',6)
      # assert_equal [[7, 1, "=SUM([.A1:.A6])"],
      #  [7, 2, "=SUM([.$A$1:.B6])"],
      #[7, 3, "=[Sheet2.A1]"],
      #[8, 2, "=SUM([.$A$1:.B7])"],
      #], oo.formulas(oo.sheets.first)
      assert_equal [[7, 1, 'SUM(A1:A6)'],
        [7, 2, 'SUM($A$1:B6)'],
        # [7, 3, "=[Sheet2.A1]"],
        # [8, 2, "=SUM([.$A$1:.B7])"],
      ], oo.formulas(oo.sheets.first)

      # setting a cell
      oo.set('A',15, 41)
      assert_equal 41, oo.cell('A',15)
      oo.set('A',16, "41")
      assert_equal "41", oo.cell('A',16)
      oo.set('A',17, 42.5)
      assert_equal 42.5, oo.cell('A',17)
    end
  end

  def test_borders_sheets
    with_each_spreadsheet(:name=>'borders') do |oo|
      oo.default_sheet = oo.sheets[1]
      assert_equal 6, oo.first_row
      assert_equal 11, oo.last_row
      assert_equal 4, oo.first_column
      assert_equal 8, oo.last_column

      oo.default_sheet = oo.sheets.first
      assert_equal 5, oo.first_row
      assert_equal 10, oo.last_row
      assert_equal 3, oo.first_column
      assert_equal 7, oo.last_column

      oo.default_sheet = oo.sheets[2]
      assert_equal 7, oo.first_row
      assert_equal 12, oo.last_row
      assert_equal 5, oo.first_column
      assert_equal 9, oo.last_column
    end
  end

  def test_only_one_sheet
    with_each_spreadsheet(:name=>'only_one_sheet') do |oo|
      assert_equal 42, oo.cell('B',4)
      assert_equal 43, oo.cell('C',4)
      assert_equal 44, oo.cell('D',4)
      oo.default_sheet = oo.sheets.first
      assert_equal 42, oo.cell('B',4)
      assert_equal 43, oo.cell('C',4)
      assert_equal 44, oo.cell('D',4)
    end
  end

  def test_bug_ric
    with_each_spreadsheet(:name=>'ric', :format=>:openoffice) do |oo|
      assert oo.empty?('A',1)
      assert oo.empty?('B',1)
      assert oo.empty?('C',1)
      assert oo.empty?('D',1)
      expected = 1
      letter = 'e'
      while letter <= 'u'
        assert_equal expected, oo.cell(letter,1)
        letter.succ!
        expected += 1
      end
      assert_equal 'J', oo.cell('v',1)
      assert_equal 'P', oo.cell('w',1)
      assert_equal 'B', oo.cell('x',1)
      assert_equal 'All', oo.cell('y',1)
      assert_equal 0, oo.cell('a',2)
      assert oo.empty?('b',2)
      assert oo.empty?('c',2)
      assert oo.empty?('d',2)
      assert_equal 'B', oo.cell('e',2)
      assert_equal 'B', oo.cell('f',2)
      assert_equal 'B', oo.cell('g',2)
      assert_equal 'B', oo.cell('h',2)
      assert_equal 'B', oo.cell('i',2)
      assert_equal 'B', oo.cell('j',2)
      assert_equal 'B', oo.cell('k',2)
      assert_equal 'B', oo.cell('l',2)
      assert_equal 'B', oo.cell('m',2)
      assert_equal 'B', oo.cell('n',2)
      assert_equal 'B', oo.cell('o',2)
      assert_equal 'B', oo.cell('p',2)
      assert_equal 'B', oo.cell('q',2)
      assert_equal 'B', oo.cell('r',2)
      assert_equal 'B', oo.cell('s',2)
      assert oo.empty?('t',2)
      assert oo.empty?('u',2)
      assert_equal 0  , oo.cell('v',2)
      assert_equal 0  , oo.cell('w',2)
      assert_equal 15 , oo.cell('x',2)
      assert_equal 15 , oo.cell('y',2)
    end
  end

  def test_mehrteilig
    with_each_spreadsheet(:name=>'Bibelbund1', :format=>:openoffice) do |oo|
      assert_equal "Tagebuch des Sekret\303\244rs.    Letzte Tagung 15./16.11.75 Schweiz", oo.cell(45,'A')
    end
  end

  # "/tmp/xxxx" darf man unter Windows nicht verwenden, weil das nicht erkannt
  # wird.
  # Besser: Methode um temporaeres Dir. portabel zu bestimmen
  def test_huge_document_to_csv
    if LONG_RUN
      with_each_spreadsheet(:name=>'Bibelbund', :format=>[
        :openoffice,
        :excelx
        # Google hier nicht, weil Google-Spreadsheets nicht so gross werden
        # duerfen
      ]) do |oo|
        Dir.mktmpdir do |tempdir|
          assert_equal "Tagebuch des Sekret\303\244rs.    Letzte Tagung 15./16.11.75 Schweiz", oo.cell(45,'A')
          assert_equal "Tagebuch des Sekret\303\244rs.  Nachrichten aus Chile", oo.cell(46,'A')
          assert_equal "Tagebuch aus Chile  Juli 1977", oo.cell(55,'A')
          assert oo.to_csv(File.join(tempdir,"Bibelbund.csv"))
          assert File.exists?(File.join(tempdir,"Bibelbund.csv"))
          assert_equal "", file_diff(File.join(TESTDIR, "Bibelbund.csv"), File.join(tempdir,"Bibelbund.csv")),
            "error in class #{oo.class}"
          #end
        end
      end
    end
  end

  def test_bug_quotes_excelx
    if LONG_RUN
      with_each_spreadsheet(:name=>'Bibelbund', :format=>[:openoffice, :excelx]) do |oo|
        oo.default_sheet = oo.sheets.first
        assert_equal 'Einflüsse der neuen Theologie in "de gereformeerde Kerken van Nederland"',
        oo.cell('a',76)
        oo.to_csv("csv#{$$}")
        assert_equal 'Einflüsse der neuen Theologie in "de gereformeerde Kerken van Nederland"',
        oo.cell('a',78)
        File.delete_if_exist("csv#{$$}")
      end
    end
  end

  def test_bug_mehrere_datum
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      oo.default_sheet = 'Sheet5'
      assert_equal :date, oo.celltype('A',4)
      assert_equal :date, oo.celltype('B',4)
      assert_equal :date, oo.celltype('C',4)
      assert_equal :date, oo.celltype('D',4)
      assert_equal :date, oo.celltype('E',4)
      assert_equal Date.new(2007,11,21), oo.cell('A',4)
      assert_equal Date.new(2007,11,21), oo.cell('B',4)
      assert_equal Date.new(2007,11,21), oo.cell('C',4)
      assert_equal Date.new(2007,11,21), oo.cell('D',4)
      assert_equal Date.new(2007,11,21), oo.cell('E',4)
      assert_equal :float, oo.celltype('A',5)
      assert_equal :float, oo.celltype('B',5)
      assert_equal :float, oo.celltype('C',5)
      assert_equal :float, oo.celltype('D',5)
      assert_equal :float, oo.celltype('E',5)
      assert_equal 42, oo.cell('A',5)
      assert_equal 42, oo.cell('B',5)
      assert_equal 42, oo.cell('C',5)
      assert_equal 42, oo.cell('D',5)
      assert_equal 42, oo.cell('E',5)
      assert_equal :string, oo.celltype('A',6)
      assert_equal :string, oo.celltype('B',6)
      assert_equal :string, oo.celltype('C',6)
      assert_equal :string, oo.celltype('D',6)
      assert_equal :string, oo.celltype('E',6)
      assert_equal "ABC", oo.cell('A',6)
      assert_equal "ABC", oo.cell('B',6)
      assert_equal "ABC", oo.cell('C',6)
      assert_equal "ABC", oo.cell('D',6)
      assert_equal "ABC", oo.cell('E',6)
    end
  end

  def test_multiple_sheets
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      2.times do
        oo.default_sheet = "Tabelle1"
        assert_equal 1, oo.cell(1,1)
        assert_equal 1, oo.cell(1,1,"Tabelle1")
        assert_equal "I am sheet 2", oo.cell('C',5,"Name of Sheet 2")
        sheetname = 'Sheet5'
        assert_equal :date, oo.celltype('A',4,sheetname)
        assert_equal :date, oo.celltype('B',4,sheetname)
        assert_equal :date, oo.celltype('C',4,sheetname)
        assert_equal :date, oo.celltype('D',4,sheetname)
        assert_equal :date, oo.celltype('E',4,sheetname)
        assert_equal Date.new(2007,11,21), oo.cell('A',4,sheetname)
        assert_equal Date.new(2007,11,21), oo.cell('B',4,sheetname)
        assert_equal Date.new(2007,11,21), oo.cell('C',4,sheetname)
        assert_equal Date.new(2007,11,21), oo.cell('D',4,sheetname)
        assert_equal Date.new(2007,11,21), oo.cell('E',4,sheetname)
        assert_equal :float, oo.celltype('A',5,sheetname)
        assert_equal :float, oo.celltype('B',5,sheetname)
        assert_equal :float, oo.celltype('C',5,sheetname)
        assert_equal :float, oo.celltype('D',5,sheetname)
        assert_equal :float, oo.celltype('E',5,sheetname)
        assert_equal 42, oo.cell('A',5,sheetname)
        assert_equal 42, oo.cell('B',5,sheetname)
        assert_equal 42, oo.cell('C',5,sheetname)
        assert_equal 42, oo.cell('D',5,sheetname)
        assert_equal 42, oo.cell('E',5,sheetname)
        assert_equal :string, oo.celltype('A',6,sheetname)
        assert_equal :string, oo.celltype('B',6,sheetname)
        assert_equal :string, oo.celltype('C',6,sheetname)
        assert_equal :string, oo.celltype('D',6,sheetname)
        assert_equal :string, oo.celltype('E',6,sheetname)
        assert_equal "ABC", oo.cell('A',6,sheetname)
        assert_equal "ABC", oo.cell('B',6,sheetname)
        assert_equal "ABC", oo.cell('C',6,sheetname)
        assert_equal "ABC", oo.cell('D',6,sheetname)
        assert_equal "ABC", oo.cell('E',6,sheetname)
        oo.reload
      end
    end
  end


  def test_bug_empty_sheet
    with_each_spreadsheet(:name=>'formula', :format=>[:openoffice, :excelx]) do |oo|
      oo.default_sheet = 'Sheet3' # is an empty sheet
      Dir.mktmpdir do |tempdir|
        oo.to_csv(File.join(tempdir,"emptysheet.csv"))
        assert_equal "", `cat #{File.join(tempdir,"emptysheet.csv")}`
      end
    end
  end

  def test_find_by_row_huge_document
    if LONG_RUN
      with_each_spreadsheet(:name=>'Bibelbund', :format=>[:openoffice, :excelx]) do |oo|
        oo.default_sheet = oo.sheets.first
        rec = oo.find 20
        assert rec
        # assert_equal "Brief aus dem Sekretariat", rec[0]
        #p rec
        assert_equal "Brief aus dem Sekretariat", rec[0]['TITEL']
        rec = oo.find 22
        assert rec
        # assert_equal "Brief aus dem Skretariat. Tagung in Amberg/Opf.",rec[0]
        assert_equal "Brief aus dem Skretariat. Tagung in Amberg/Opf.",rec[0]['TITEL']
      end
    end
  end

  def test_find_by_row
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      oo.header_line = nil
      rec = oo.find 16
      assert rec
      assert_nil oo.header_line
      # keine Headerlines in diesem Beispiel definiert
      assert_equal "einundvierzig", rec[0]
      #assert_equal false, rec
      rec = oo.find 15
      assert rec
      assert_equal 41,rec[0]
    end
  end

  def test_find_by_row_if_header_line_is_not_nil
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      oo.header_line = 2
      refute_nil oo.header_line
      rec = oo.find 1
      assert rec
      assert_equal 5, rec[0]
      assert_equal 6, rec[1]
      rec = oo.find 15
      assert rec
      assert_equal "einundvierzig", rec[0]
    end
  end

  def test_find_by_conditions
    if LONG_RUN
      with_each_spreadsheet(:name=>'Bibelbund', :format=>[:openoffice,
          :excelx]) do |oo|
        #-----------------------------------------------------------------
        zeilen = oo.find(:all, :conditions => {
            'TITEL' => 'Brief aus dem Sekretariat'
          }
        )
        assert_equal 2, zeilen.size
        assert_equal [{"VERFASSER"=>"Almassy, Annelene von",
            "INTERNET"=>nil,
            "SEITE"=>316.0,
            "KENNUNG"=>"Aus dem Bibelbund",
            "OBJEKT"=>"Bibel+Gem",
            "PC"=>"#C:\\Bibelbund\\reprint\\BuG1982-3.pdf#",
            "NUMMER"=>"1982-3",
            "TITEL"=>"Brief aus dem Sekretariat"},
          {"VERFASSER"=>"Almassy, Annelene von",
            "INTERNET"=>nil,
            "SEITE"=>222.0,
            "KENNUNG"=>"Aus dem Bibelbund",
            "OBJEKT"=>"Bibel+Gem",
            "PC"=>"#C:\\Bibelbund\\reprint\\BuG1983-2.pdf#",
            "NUMMER"=>"1983-2",
            "TITEL"=>"Brief aus dem Sekretariat"}] , zeilen

        #----------------------------------------------------------
        zeilen = oo.find(:all,
          :conditions => { 'VERFASSER' => 'Almassy, Annelene von' }
        )
        assert_equal 13, zeilen.size
        #----------------------------------------------------------
        zeilen = oo.find(:all, :conditions => {
            'TITEL' => 'Brief aus dem Sekretariat',
            'VERFASSER' => 'Almassy, Annelene von',
          }
        )
        assert_equal 2, zeilen.size
        assert_equal [{"VERFASSER"=>"Almassy, Annelene von",
            "INTERNET"=>nil,
            "SEITE"=>316.0,
            "KENNUNG"=>"Aus dem Bibelbund",
            "OBJEKT"=>"Bibel+Gem",
            "PC"=>"#C:\\Bibelbund\\reprint\\BuG1982-3.pdf#",
            "NUMMER"=>"1982-3",
            "TITEL"=>"Brief aus dem Sekretariat"},
          {"VERFASSER"=>"Almassy, Annelene von",
            "INTERNET"=>nil,
            "SEITE"=>222.0,
            "KENNUNG"=>"Aus dem Bibelbund",
            "OBJEKT"=>"Bibel+Gem",
            "PC"=>"#C:\\Bibelbund\\reprint\\BuG1983-2.pdf#",
            "NUMMER"=>"1983-2",
            "TITEL"=>"Brief aus dem Sekretariat"}] , zeilen

        # Result as an array
        zeilen = oo.find(:all,
          :conditions => {
            'TITEL' => 'Brief aus dem Sekretariat',
            'VERFASSER' => 'Almassy, Annelene von',
          }, :array => true)
        assert_equal 2, zeilen.size
        assert_equal [
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
          ]] , zeilen
      end
    end
  end


  #TODO: temporaerer Test
  def test_seiten_als_date
    if LONG_RUN
      with_each_spreadsheet(:name=>'Bibelbund', :format=>:excelx) do |oo|
        assert_equal 'Bericht aus dem Sekretariat', oo.cell(13,1)
        assert_equal '1981-4', oo.cell(13,'D')
        assert_equal String, oo.excelx_type(13,'E')[1].class
        assert_equal [:numeric_or_formula,"General"], oo.excelx_type(13,'E')
        assert_equal '428', oo.excelx_value(13,'E')
        assert_equal 428.0, oo.cell(13,'E')
      end
    end
  end

  def test_column
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      expected = [1.0,5.0,nil,10.0,Date.new(1961,11,21),'tata',nil,nil,nil,nil,'thisisa11',41.0,nil,nil,41.0,'einundvierzig',nil,Date.new(2007,5,31)]
      assert_equal expected, oo.column(1)
      assert_equal expected, oo.column('a')
    end
  end

  def test_column_huge_document
    if LONG_RUN
      with_each_spreadsheet(:name=>'Bibelbund', :format=>[:openoffice,
          :excelx]) do |oo|
        oo.default_sheet = oo.sheets.first
        assert_equal 3735, oo.column('a').size
        #assert_equal 499, oo.column('a').size
      end
    end
  end

  def test_simple_spreadsheet_find_by_condition
    with_each_spreadsheet(:name=>'simple_spreadsheet') do |oo|
      oo.header_line = 3
      # oo.date_format = '%m/%d/%Y' if oo.class == Google
      erg = oo.find(:all, :conditions => {'Comment' => 'Task 1'})
      assert_equal Date.new(2007,05,07), erg[1]['Date']
      assert_equal 10.75       , erg[1]['Start time']
      assert_equal 12.50       , erg[1]['End time']
      assert_equal 0           , erg[1]['Pause']
      assert_equal 1.75        , erg[1]['Sum']
      assert_equal "Task 1"    , erg[1]['Comment']
    end
  end

  def get_extension(oo)
    case oo
    when Roo::OpenOffice
      ".ods"
    when Roo::Excelx
      ".xlsx"
    end
  end

  def test_info
    expected_templ = "File: numbers1%s\n"+
      "Number of sheets: 5\n"+
      "Sheets: Tabelle1, Name of Sheet 2, Sheet3, Sheet4, Sheet5\n"+
      "Sheet 1:\n"+
      "  First row: 1\n"+
      "  Last row: 18\n"+
      "  First column: A\n"+
      "  Last column: G\n"+
      "Sheet 2:\n"+
      "  First row: 5\n"+
      "  Last row: 14\n"+
      "  First column: B\n"+
      "  Last column: E\n"+
      "Sheet 3:\n"+
      "  First row: 1\n"+
      "  Last row: 1\n"+
      "  First column: A\n"+
      "  Last column: BA\n"+
      "Sheet 4:\n"+
      "  First row: 1\n"+
      "  Last row: 1\n"+
      "  First column: A\n"+
      "  Last column: E\n"+
      "Sheet 5:\n"+
      "  First row: 1\n"+
      "  Last row: 6\n"+
      "  First column: A\n"+
      "  Last column: E"
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      ext = get_extension(oo)
      expected = sprintf(expected_templ,ext)
      begin
        if oo.class == Google
          assert_equal expected.gsub(/numbers1/,key_of("numbers1")), oo.info
        else
          assert_equal expected, oo.info
        end
      rescue NameError
        #
      end
    end
  end

  def test_info_doesnt_set_default_sheet
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      oo.default_sheet = 'Sheet3'
      oo.info
      assert_equal 'Sheet3', oo.default_sheet
    end
  end

  def test_bug_bbu
    with_each_spreadsheet(:name=>'bbu', :format=>[:openoffice, :excelx]) do |oo|
        assert_equal "File: bbu#{get_extension(oo)}
Number of sheets: 3
Sheets: 2007_12, Tabelle2, Tabelle3
Sheet 1:
  First row: 1
  Last row: 4
  First column: A
  Last column: F
Sheet 2:
  - empty -
Sheet 3:
  - empty -", oo.info

      oo.default_sheet = oo.sheets[1] # empty sheet
      assert_nil oo.first_row
      assert_nil oo.last_row
      assert_nil oo.first_column
      assert_nil oo.last_column
    end
  end


  def test_bug_time_nil
    with_each_spreadsheet(:name=>'time-test') do |oo|
      assert_equal 12*3600+13*60+14, oo.cell('B',1) # 12:13:14 (secs since midnight)
      assert_equal :time, oo.celltype('B',1)
      assert_equal 15*3600+16*60, oo.cell('C',1) # 15:16    (secs since midnight)
      assert_equal :time, oo.celltype('C',1)
      assert_equal 23*3600, oo.cell('D',1) # 23:00    (secs since midnight)
      assert_equal :time, oo.celltype('D',1)
    end
  end

  def test_date_time_to_csv
    with_each_spreadsheet(:name=>'time-test') do |oo|
      Dir.mktmpdir do |tempdir|
        csv_output = File.join(tempdir,'time_test.csv')
        assert oo.to_csv(csv_output)
        assert File.exists?(csv_output)
        assert_equal "", `diff --strip-trailing-cr #{TESTDIR}/time-test.csv #{csv_output}`
        # --strip-trailing-cr is needed because the test-file use 0A and
        # the test on an windows box generates 0D 0A as line endings
      end
    end
  end

  def test_boolean_to_csv
    with_each_spreadsheet(:name=>'boolean') do |oo|
      Dir.mktmpdir do |tempdir|
        csv_output = File.join(tempdir,'boolean.csv')
        assert oo.to_csv(csv_output)
        assert File.exists?(csv_output)
        assert_equal "", `diff --strip-trailing-cr #{TESTDIR}/boolean.csv #{csv_output}`
        # --strip-trailing-cr is needed because the test-file use 0A and
        # the test on an windows box generates 0D 0A as line endings
      end
    end
  end
  def test_link_to_csv
    with_each_spreadsheet(:name=>'link',:format=>:excelx) do |oo|
      Dir.mktmpdir do |tempdir|
        csv_output = File.join(tempdir,'link.csv')
        assert oo.to_csv(csv_output)
        assert File.exists?(csv_output)
        assert_equal "", `diff --strip-trailing-cr #{TESTDIR}/link.csv #{csv_output}`
        # --strip-trailing-cr is needed because the test-file use 0A and
        # the test on an windows box generates 0D 0A as line endings
      end
    end
  end
  def test_date_time_yaml
    with_each_spreadsheet(:name=>'time-test') do |oo|
      expected =
        "--- \ncell_1_1: \n  row: 1 \n  col: 1 \n  celltype: string \n  value: Mittags: \ncell_1_2: \n  row: 1 \n  col: 2 \n  celltype: time \n  value: 12:13:14 \ncell_1_3: \n  row: 1 \n  col: 3 \n  celltype: time \n  value: 15:16:00 \ncell_1_4: \n  row: 1 \n  col: 4 \n  celltype: time \n  value: 23:00:00 \ncell_2_1: \n  row: 2 \n  col: 1 \n  celltype: date \n  value: 2007-11-21 \n"
      assert_equal expected, oo.to_yaml
    end
  end

  # Erstellt eine Liste aller Zellen im Spreadsheet. Dies ist nötig, weil ein einfacher
  # Textvergleich des XML-Outputs nicht funktioniert, da xml-builder die Attribute
  # nicht immer in der gleichen Reihenfolge erzeugt.
  def init_all_cells(oo,sheet)
    all = []
    oo.first_row(sheet).upto(oo.last_row(sheet)) do |row|
      oo.first_column(sheet).upto(oo.last_column(sheet)) do |col|
        unless oo.empty?(row,col,sheet)
          all << {:row => row.to_s,
            :column => col.to_s,
            :content => oo.cell(row,col,sheet).to_s,
            :type => oo.celltype(row,col,sheet).to_s,
          }
        end
      end
    end
    all
  end

  def test_to_xml
    with_each_spreadsheet(:name=>'numbers1', :encoding => 'utf8') do |oo|
      skip if defined? JRUBY_VERSION
      oo.to_xml
      sheetname = oo.sheets.first
      doc = Nokogiri::XML(oo.to_xml)
      sheet_count = 0
      doc.xpath('//spreadsheet/sheet').each {|tmpelem|
        sheet_count += 1
      }
      assert_equal 5, sheet_count
      doc.xpath('//spreadsheet/sheet').each { |xml_sheet|
        all_cells = init_all_cells(oo, sheetname)
        x = 0
        assert_equal sheetname, xml_sheet.attributes['name'].value
        xml_sheet.children.each {|cell|
          if cell.attributes['name']
            expected = [all_cells[x][:row],
              all_cells[x][:column],
              all_cells[x][:content],
              all_cells[x][:type],
            ]
            result = [
              cell.attributes['row'],
              cell.attributes['column'],
              cell.content,
              cell.attributes['type'],
            ]
            assert_equal expected, result
            x += 1
          end # if
        } # end of sheet
        sheetname = oo.sheets[oo.sheets.index(sheetname)+1]
      }
    end
  end

  def test_bug_to_xml_with_empty_sheets
    with_each_spreadsheet(:name=>'emptysheets', :format=>[:openoffice, :excelx]) do |oo|
      oo.sheets.each { |sheet|
        assert_nil oo.first_row, "first_row not nil in sheet #{sheet}"
        assert_nil oo.last_row, "last_row not nil in sheet #{sheet}"
        assert_nil oo.first_column, "first_column not nil in sheet #{sheet}"
        assert_nil oo.last_column, "last_column not nil in sheet #{sheet}"
        assert_nil oo.first_row(sheet), "first_row not nil in sheet #{sheet}"
        assert_nil oo.last_row(sheet), "last_row not nil in sheet #{sheet}"
        assert_nil oo.first_column(sheet), "first_column not nil in sheet #{sheet}"
        assert_nil oo.last_column(sheet), "last_column not nil in sheet #{sheet}"
      }
      oo.to_xml
    end
  end

  def test_bug_simple_spreadsheet_time_bug
    # really a bug? are cells really of type time?
    # No! :float must be the correct type
    with_each_spreadsheet(:name=>'simple_spreadsheet', :format=>:excelx) do |oo|
      # puts oo.cell('B',5).to_s
      # assert_equal :time, oo.celltype('B',5)
      assert_equal :float, oo.celltype('B',5)
      assert_equal 10.75, oo.cell('B',5)
      assert_equal 12.50, oo.cell('C',5)
      assert_equal 0, oo.cell('D',5)
      assert_equal 1.75, oo.cell('E',5)
      assert_equal 'Task 1', oo.cell('F',5)
      assert_equal Date.new(2007,5,7), oo.cell('A',5)
    end
  end

  def test_simple2_excelx
    with_each_spreadsheet(:name=>'simple_spreadsheet', :format=>:excelx) do |oo|
      assert_equal [:numeric_or_formula, "yyyy\\-mm\\-dd"], oo.excelx_type('A',4)
      assert_equal [:numeric_or_formula, "#,##0.00"], oo.excelx_type('B',4)
      assert_equal [:numeric_or_formula, "#,##0.00"], oo.excelx_type('c',4)
      assert_equal [:numeric_or_formula, "General"], oo.excelx_type('d',4)
      assert_equal [:numeric_or_formula, "General"], oo.excelx_type('e',4)
      assert_equal :string, oo.excelx_type('f',4)

      assert_equal "39209", oo.excelx_value('a',4)
      assert_equal "yyyy\\-mm\\-dd", oo.excelx_format('a',4)
      assert_equal "9.25", oo.excelx_value('b',4)
      assert_equal "10.25", oo.excelx_value('c',4)
      assert_equal "0", oo.excelx_value('d',4)
      #... Sum-Spalte
      # assert_equal "Task 1", oo.excelx_value('f',4)
      assert_equal "Task 1", oo.cell('f',4)
      assert_equal Date.new(2007,05,07), oo.cell('a',4)
      assert_equal "9.25", oo.excelx_value('b',4)
      assert_equal "#,##0.00", oo.excelx_format('b',4)
      assert_equal 9.25, oo.cell('b',4)
      assert_equal :float, oo.celltype('b',4)
      assert_equal :float, oo.celltype('d',4)
      assert_equal 0, oo.cell('d',4)
      assert_equal :formula, oo.celltype('e',4)
      assert_equal 1, oo.cell('e',4)
      assert_equal 'C4-B4-D4', oo.formula('e',4)
      assert_equal :string, oo.celltype('f',4)
      assert_equal "Task 1", oo.cell('f',4)
    end
  end

  def test_datetime
    with_each_spreadsheet(:name=>'datetime') do |oo|
      val = oo.cell('c',3)
      assert_equal :datetime, oo.celltype('c',3)
      assert_equal DateTime.new(1961,11,21,12,17,18), val
      assert_kind_of DateTime, val
      val = oo.cell('a',1)
      assert_equal :date, oo.celltype('a',1)
      assert_kind_of Date, val
      assert_equal Date.new(1961,11,21), val
      assert_equal Date.new(1961,11,21), oo.cell('a',1)
      assert_equal DateTime.new(1961,11,21,12,17,18), oo.cell('a',3)
      assert_equal DateTime.new(1961,11,21,12,17,18), oo.cell('b',3)
      assert_equal DateTime.new(1961,11,21,12,17,18), oo.cell('c',3)
      assert_equal DateTime.new(1961,11,21,12,17,18), oo.cell('a',4)
      assert_equal DateTime.new(1961,11,21,12,17,18), oo.cell('b',4)
      assert_equal DateTime.new(1961,11,21,12,17,18), oo.cell('c',4)
      assert_equal DateTime.new(1961,11,21,12,17,18), oo.cell('a',5)
      assert_equal DateTime.new(1961,11,21,12,17,18), oo.cell('b',5)
      assert_equal DateTime.new(1961,11,21,12,17,18), oo.cell('c',5)
      assert_equal Date.new(1961,11,21), oo.cell('a',6)
      assert_equal Date.new(1961,11,21), oo.cell('b',6)
      assert_equal Date.new(1961,11,21), oo.cell('c',6)
      assert_equal Date.new(1961,11,21), oo.cell('a',7)
      assert_equal Date.new(1961,11,21), oo.cell('b',7)
      assert_equal Date.new(1961,11,21), oo.cell('c',7)
      assert_equal DateTime.new(2013,11,5,11,45,00), oo.cell('a',8)
      assert_equal DateTime.new(2013,11,5,11,45,00), oo.cell('b',8)
      assert_equal DateTime.new(2013,11,5,11,45,00), oo.cell('c',8)
    end
  end

  def test_cell_openoffice_html_escape
    with_each_spreadsheet(:name=>'html-escape', :format=>:openoffice) do |oo|
      assert_equal "'", oo.cell(1,1)
      assert_equal "&", oo.cell(2,1)
      assert_equal ">", oo.cell(3,1)
      assert_equal "<", oo.cell(4,1)
      assert_equal "`", oo.cell(5,1)
      # test_openoffice_zipped will catch issues with &quot;
    end
  end

  def test_cell_boolean
    with_each_spreadsheet(:name=>'boolean', :format=>[:openoffice, :excelx]) do |oo|
      if oo.class == Roo::Excelx
        assert_equal true, oo.cell(1, 1), "failure in #{oo.class}"
        assert_equal false, oo.cell(2, 1), "failure in #{oo.class}"

        cell = oo.sheet_for(oo.default_sheet).cells[[1, 1,]]
        assert_equal 'TRUE', cell.formatted_value

        cell = oo.sheet_for(oo.default_sheet).cells[[2, 1,]]
        assert_equal 'FALSE', cell.formatted_value
      else
        assert_equal "true", oo.cell(1,1), "failure in "+oo.class.to_s
        assert_equal "false", oo.cell(2,1), "failure in "+oo.class.to_s
      end
    end
  end

  def test_cell_multiline
    with_each_spreadsheet(:name=>'paragraph', :format=>[:openoffice, :excelx]) do |oo|
      assert_equal "This is a test\nof a multiline\nCell", oo.cell(1,1)
      assert_equal "This is a test\n¶\nof a multiline\n\nCell", oo.cell(1,2)
      assert_equal "first p\n\nsecond p\n\nlast p", oo.cell(2,1)
    end
  end

  def test_cell_styles
    # styles only valid in excel spreadsheets?
    # TODO: what todo with other spreadsheet types
    with_each_spreadsheet(:name=>'style', :format=>[# :openoffice,
        :excelx
      ]) do |oo|
      # bold
      assert_equal true,  oo.font(1,1).bold?
      assert_equal false, oo.font(1,1).italic?
      assert_equal false, oo.font(1,1).underline?

      # italic
      assert_equal false, oo.font(2,1).bold?
      assert_equal true,  oo.font(2,1).italic?
      assert_equal false, oo.font(2,1).underline?

      # normal
      assert_equal false, oo.font(3,1).bold?
      assert_equal false, oo.font(3,1).italic?
      assert_equal false, oo.font(3,1).underline?

      # underline
      assert_equal false, oo.font(4,1).bold?
      assert_equal false, oo.font(4,1).italic?
      assert_equal true,  oo.font(4,1).underline?

      # bold italic
      assert_equal true,  oo.font(5,1).bold?
      assert_equal true,  oo.font(5,1).italic?
      assert_equal false, oo.font(5,1).underline?

      # bold underline
      assert_equal true,  oo.font(6,1).bold?
      assert_equal false, oo.font(6,1).italic?
      assert_equal true,  oo.font(6,1).underline?

      # italic underline
      assert_equal false, oo.font(7,1).bold?
      assert_equal true,  oo.font(7,1).italic?
      assert_equal true,  oo.font(7,1).underline?

      # bolded row
      assert_equal true, oo.font(8,1).bold?
      assert_equal false,  oo.font(8,1).italic?
      assert_equal false,  oo.font(8,1).underline?

      # bolded col
      assert_equal true, oo.font(9,2).bold?
      assert_equal false,  oo.font(9,2).italic?
      assert_equal false,  oo.font(9,2).underline?

      # bolded row, italic col
      assert_equal true, oo.font(10,3).bold?
      assert_equal true,  oo.font(10,3).italic?
      assert_equal false,  oo.font(10,3).underline?

      # normal
      assert_equal false, oo.font(11,4).bold?
      assert_equal false,  oo.font(11,4).italic?
      assert_equal false,  oo.font(11,4).underline?
    end
  end

  # Need to extend to other formats
  def test_row_whitespace
    # auf dieses Dokument habe ich keinen Zugriff TODO:
    with_each_spreadsheet(:name=>'whitespace') do |oo|
      oo.default_sheet = "Sheet1"
      assert_equal [nil, nil, nil, nil, nil, nil], oo.row(1)
      assert_equal [nil, nil, nil, nil, nil, nil], oo.row(2)
      assert_equal ["Date", "Start time", "End time", "Pause", "Sum", "Comment"], oo.row(3)
      assert_equal [Date.new(2007,5,7), 9.25, 10.25, 0.0, 1.0, "Task 1"], oo.row(4)
      assert_equal [nil, nil, nil, nil, nil, nil], oo.row(5)
      assert_equal [Date.new(2007,5,7), 10.75, 10.75, 0.0, 0.0, "Task 1"], oo.row(6)
      oo.default_sheet = "Sheet2"
      assert_equal ["Date", nil, "Start time"], oo.row(1)
      assert_equal [Date.new(2007,5,7), nil, 9.25], oo.row(2)
      assert_equal [Date.new(2007,5,7), nil,  10.75], oo.row(3)
    end
  end

  def test_col_whitespace
    #TODO:
    # kein Zugriff auf Dokument whitespace
    with_each_spreadsheet(:name=>'whitespace') do |oo|
      oo.default_sheet = "Sheet1"
      assert_equal ["Date", Date.new(2007,5,7), nil, Date.new(2007,5,7)], oo.column(1)
      assert_equal ["Start time", 9.25, nil, 10.75], oo.column(2)
      assert_equal ["End time", 10.25, nil, 10.75], oo.column(3)
      assert_equal ["Pause", 0.0, nil, 0.0], oo.column(4)
      assert_equal ["Sum", 1.0, nil, 0.0], oo.column(5)
      assert_equal ["Comment","Task 1", nil, "Task 1"], oo.column(6)
      oo.default_sheet = "Sheet2"
      assert_equal [nil, nil, nil], oo.column(1)
      assert_equal [nil, nil, nil], oo.column(2)
      assert_equal ["Date", Date.new(2007,5,7), Date.new(2007,5,7)], oo.column(3)
      assert_equal [nil, nil, nil], oo.column(4)
      assert_equal [ "Start time", 9.25, 10.75], oo.column(5)
    end
  end

  def test_excelx_links
    with_each_spreadsheet(:name=>'link', :format=>:excelx) do |oo|
      assert_equal 'Google', oo.cell(1,1)
      assert_equal 'http://www.google.com', oo.cell(1,1).href
    end
  end

  # Excel has two base date formats one from 1900 and the other from 1904.
  # see #test_base_dates_in_excel
  def test_base_dates_in_excelx
    with_each_spreadsheet(:name=>'1900_base', :format=>:excelx) do |oo|
      assert_equal Date.new(2009,06,15), oo.cell(1,1)
      assert_equal :date, oo.celltype(1,1)
    end
    with_each_spreadsheet(:name=>'1904_base', :format=>:excelx) do |oo|
      assert_equal Date.new(2009,06,15), oo.cell(1,1)
      assert_equal :date, oo.celltype(1,1)
    end
  end

  def test_cell_methods
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      assert_equal 10, oo.a4 # cell(4,'A')
      assert_equal 11, oo.b4 # cell(4,'B')
      assert_equal 12, oo.c4 # cell(4,'C')
      assert_equal 13, oo.d4 # cell(4,'D')
      assert_equal 14, oo.e4 # cell(4,'E')
      assert_equal 'ABC', oo.c6('Sheet5')
      assert_equal 41, oo.a12

      assert_raises(NoMethodError) do
        # a42a is not a valid cell name, should raise ArgumentError
        assert_equal 9999, oo.a42a
      end
    end
  end


  # compare large spreadsheets
  def test_compare_large_spreadsheets
    # problematisch, weil Formeln in Excel nicht unterstützt werden
    if LONG_RUN
      qq = Roo::OpenOffice.new(File.join('test',"Bibelbund.ods"))
      with_each_spreadsheet(:name=>'Bibelbund') do |oo|
        # p "comparing Bibelbund.ods with #{oo.class}"
        oo.sheets.each do |sh|
          oo.first_row.upto(oo.last_row) do |row|
            oo.first_column.upto(oo.last_column) do |col|
              c1 = qq.cell(row,col,sh)
              c1.force_encoding("UTF-8") if c1.class == String
              c2 = oo.cell(row,col,sh)
              c2.force_encoding("UTF-8") if c2.class == String
              assert_equal c1, c2, "diff in #{sh}/#{row}/#{col}}"
              assert_equal qq.celltype(row,col,sh), oo.celltype(row,col,sh)
              assert_equal qq.formula?(row,col,sh), oo.formula?(row,col,sh) if oo.class != Roo::Excel
            end
          end
        end
      end
    end # LONG_RUN
  end

  def test_label
    with_each_spreadsheet(:name=>'named_cells', :format=>[:openoffice,:excelx,:libreoffice]) do |oo|
      # oo.default_sheet = oo.sheets.first
      begin
        row,col = oo.label('anton')
      rescue ArgumentError
        puts "labels error at #{oo.class}"
        raise
      end
      assert_equal 5, row, "error with label in class #{oo.class}"
      assert_equal 3, col, "error with label in class #{oo.class}"

      row,col = oo.label('anton')
      assert_equal 'Anton', oo.cell(row,col), "error with label in class #{oo.class}"

      row,col = oo.label('berta')
      assert_equal 'Bertha', oo.cell(row,col), "error with label in class #{oo.class}"

      row,col = oo.label('caesar')
      assert_equal 'Cäsar', oo.cell(row,col),"error with label in class #{oo.class}"

      row,col = oo.label('never')
      assert_nil row
      assert_nil col

      row,col,sheet = oo.label('anton')
      assert_equal 5, row
      assert_equal 3, col
      assert_equal "Sheet1", sheet
    end
  end

  def test_method_missing_anton
    with_each_spreadsheet(:name=>'named_cells', :format=>[:openoffice,:excelx,:libreoffice]) do |oo|
      # oo.default_sheet = oo.sheets.first
      assert_equal "Anton", oo.anton
      assert_raises(NoMethodError) {
        oo.never
      }
    end
  end

  def test_labels
    with_each_spreadsheet(:name=>'named_cells', :format=>[:openoffice,:excelx,:libreoffice]) do |oo|
      # oo.default_sheet = oo.sheets.first
      assert_equal [
        ['anton',[5,3,'Sheet1']],
        ['berta',[4,2,'Sheet1']],
        ['caesar',[7,2,'Sheet1']],
      ], oo.labels, "error with labels array in class #{oo.class}"
    end
  end

   def test_labeled_cells
     with_each_spreadsheet(:name=>'named_cells', :format=>[:openoffice,:excelx,:libreoffice]) do |oo|
       oo.default_sheet = oo.sheets.first
       begin
         row,col = oo.label('anton')
       rescue ArgumentError
         puts "labels error at #{oo.class}"
         raise
       end
       assert_equal 5, row
       assert_equal 3, col

       row,col = oo.label('anton')
       assert_equal 'Anton', oo.cell(row,col)

       row,col = oo.label('berta')
       assert_equal 'Bertha', oo.cell(row,col)

       row,col = oo.label('caesar')
       assert_equal 'Cäsar', oo.cell(row,col)

       row,col = oo.label('never')
       assert_nil row
       assert_nil col

       row,col,sheet = oo.label('anton')
       assert_equal 5, row
       assert_equal 3, col
       assert_equal "Sheet1", sheet

       assert_equal "Anton", oo.anton
       assert_raises(NoMethodError) {
         row,col = oo.never
       }

  # Reihenfolge row,col,sheet analog zu #label
       assert_equal [
         ['anton',[5,3,'Sheet1']],
         ['berta',[4,2,'Sheet1']],
         ['caesar',[7,2,'Sheet1']],
       ], oo.labels, "error with labels array in class #{oo.class}"
     end
   end

  require 'matrix'
  def test_matrix
    with_each_spreadsheet(:name => 'matrix', :format => :openoffice) do |oo|
      oo.default_sheet = oo.sheets.first
      assert_equal Matrix[
        [1.0, 2.0, 3.0],
        [4.0, 5.0, 6.0],
        [7.0, 8.0, 9.0] ], oo.to_matrix
    end
  end

  def test_matrix_selected_range
    with_each_spreadsheet(:name => 'matrix', :format=>:openoffice) do |oo|
      oo.default_sheet = 'Sheet2'
      assert_equal Matrix[
        [1.0, 2.0, 3.0],
        [4.0, 5.0, 6.0],
        [7.0, 8.0, 9.0] ], oo.to_matrix(3,4,5,6)
    end
  end

  def test_matrix_all_nil
    with_each_spreadsheet(:name => 'matrix', :format=>:openoffice) do |oo|
      oo.default_sheet = 'Sheet2'
      assert_equal Matrix[
        [nil, nil, nil],
        [nil, nil, nil],
        [nil, nil, nil] ], oo.to_matrix(10,10,12,12)
    end
  end

  def test_matrix_values_and_nil
    with_each_spreadsheet(:name => 'matrix', :format=>:openoffice) do |oo|
      oo.default_sheet = 'Sheet3'
      assert_equal Matrix[
        [1.0, nil, 3.0],
        [4.0, 5.0, 6.0],
        [7.0, 8.0, nil] ], oo.to_matrix(1,1,3,3)
    end
  end

  def test_matrix_specifying_sheet
    with_each_spreadsheet(:name => 'matrix', :format => :openoffice) do |oo|
      oo.default_sheet = oo.sheets.first
      assert_equal Matrix[
        [1.0, nil, 3.0],
        [4.0, 5.0, 6.0],
        [7.0, 8.0, nil] ], oo.to_matrix(nil, nil, nil, nil, 'Sheet3')
    end
  end

  # #formulas of an empty sheet should return an empty array and not result in
  # an error message
  # 2011-06-24
  def test_bug_formulas_empty_sheet
    with_each_spreadsheet(:name =>'emptysheets',
      :format=>[:openoffice,:excelx]) do |oo|
        oo.default_sheet = oo.sheets.first
        oo.formulas
      assert_equal([], oo.formulas)
    end
  end

  # #to_yaml of an empty sheet should return an empty string and not result in
  # an error message
  # 2011-06-24
  def test_bug_to_yaml_empty_sheet
    with_each_spreadsheet(:name =>'emptysheets',
      :format=>[:openoffice,:excelx]) do |oo|
        oo.default_sheet = oo.sheets.first
        oo.to_yaml
      assert_equal('', oo.to_yaml)
    end
  end

  # #to_matrix of an empty sheet should return an empty matrix and not result in
  # an error message
  # 2011-06-25
  def test_bug_to_matrix_empty_sheet
    with_each_spreadsheet(:name =>'emptysheets',
      :format=>[:openoffice,:excelx]) do |oo|
        oo.default_sheet = oo.sheets.first
        oo.to_matrix
      assert_equal(Matrix.empty(0,0), oo.to_matrix)
    end
  end

  # 2011-08-03
  def test_bug_datetime_to_csv
    with_each_spreadsheet(:name=>'datetime') do |oo|
      Dir.mktmpdir do |tempdir|
        datetime_csv_file = File.join(tempdir,"datetime.csv")

        assert oo.to_csv(datetime_csv_file)
        assert File.exists?(datetime_csv_file)
        assert_equal "", file_diff('test/files/so_datetime.csv', datetime_csv_file)
      end
    end
  end

  def test_bug_pfand_from_windows_phone_xlsx
    return if defined? JRUBY_VERSION
    with_each_spreadsheet(:name=>'Pfand_from_windows_phone', :format=>:excelx) do |oo|
      oo.default_sheet = oo.sheets.first
      assert_equal ['Blatt1','Blatt2','Blatt3'], oo.sheets
      assert_equal 'Summe', oo.cell('b',1)

      assert_equal Date.new(2011,9,14), oo.cell('a',2)
      assert_equal :date, oo.celltype('a',2)
      assert_equal Date.new(2011,9,15), oo.cell('a',3)
      assert_equal :date, oo.celltype('a',3)

      assert_equal 3.81, oo.cell('b',2)
      assert_equal "SUM(C2:L2)", oo.formula('b',2)
      assert_equal 0.7, oo.cell('c',2)
    end # each
  end

  def test_comment
    with_each_spreadsheet(:name=>'comments', :format=>[:openoffice,:libreoffice,
        :excelx]) do |oo|
      oo.default_sheet = oo.sheets.first
      assert_equal 'Kommentar fuer B4',oo.comment('b',4)
      assert_equal 'Kommentar fuer B5',oo.comment('b',5)
      assert_nil oo.comment('b',99)
      # no comment at the second page
      oo.default_sheet = oo.sheets[1]
      assert_nil oo.comment('b',4)
    end
  end

  def test_comments
    with_each_spreadsheet(:name=>'comments', :format=>[:openoffice,:libreoffice,
        :excelx]) do |oo|
      oo.default_sheet = oo.sheets.first
      assert_equal [
        [4, 2, "Kommentar fuer B4"],
        [5, 2, "Kommentar fuer B5"],
      ], oo.comments(oo.sheets.first), "comments error in class #{oo.class}"
      # no comments at the second page
      oo.default_sheet = oo.sheets[1]
      assert_equal [], oo.comments, "comments error in class #{oo.class}"
    end

    with_each_spreadsheet(:name=>'comments-google', :format=>[:excelx]) do |oo|
      oo.default_sheet = oo.sheets.first
      assert_equal [[1, 1, "this is a comment\n\t-Steven Daniels"]], oo.comments(oo.sheets.first), "comments error in class #{oo.class}"
    end
  end

  def common_possible_bug_snowboard_cells(ss)
    assert_equal "A.", ss.cell(13,'A'), ss.class
    assert_equal 147, ss.cell(13,'f'), ss.class
    assert_equal 152, ss.cell(13,'g'), ss.class
    assert_equal 156, ss.cell(13,'h'), ss.class
    assert_equal 158, ss.cell(13,'i'), ss.class
    assert_equal 160, ss.cell(13,'j'), ss.class
    assert_equal 164, ss.cell(13,'k'), ss.class
    assert_equal 168, ss.cell(13,'l'), ss.class
    assert_equal :string, ss.celltype(13,'m'), ss.class
    assert_equal "159W", ss.cell(13,'m'), ss.class
    assert_equal "164W", ss.cell(13,'n'), ss.class
    assert_equal "168W", ss.cell(13,'o'), ss.class
  end

  def split_coord(s)
    letter = ""
    number = 0
    i = 0
    while i<s.length and "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz".include?(s[i,1])
      letter += s[i,1]
      i+=1
    end
    while i<s.length and "01234567890".include?(s[i,1])
      number = number*10 + s[i,1].to_i
      i+=1
    end
    if letter=="" or number==0
      raise ArgumentError
    end
    return letter,number
  end

  def test_bug_numbered_sheet_names
    with_each_spreadsheet(:name=>'bug-numbered-sheet-names', :format=>:excelx) do |oo|
      oo.each_with_pagename { }
    end
  end

  def test_close
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      next unless (tempdir = oo.instance_variable_get('@tmpdir'))
      oo.close
      assert !File.exists?(tempdir), "Expected #{tempdir} to be cleaned up, but it still exists"
    end
  end

  # NOTE: Ruby 2.4.0 changed the way GC works. The last Roo object created by
  #       with_each_spreadsheet wasn't getting GC'd until after the process
  #       ended.
  #
  #       That behavior change broke this test. In order to fix it, I forked the
  #       process and passed the temp directories from the forked process in
  #       order to check if they were removed properly.
  def test_finalize
    read, write = IO.pipe
    Process.fork do
      with_each_spreadsheet(name: "numbers1") do |oo|
        write.puts oo.instance_variable_get("@tmpdir")
      end
    end

    Process.wait
    write.close
    tempdirs = read.read.split("\n")
    read.close

    refute tempdirs.empty?
    tempdirs.each do |tempdir|
      refute File.exist?(tempdir), "Expected #{tempdir} to be cleaned up, but it still exists"
    end
  end

  def test_cleanup_on_error
    old_temp_files = Dir.open(Dir.tmpdir).to_a
    with_each_spreadsheet(:name=>'non_existent_file', :ignore_errors=>true) do |oo|; end
    assert_equal Dir.open(Dir.tmpdir).to_a, old_temp_files
  end

  def test_name_with_leading_slash
    xlsx = Roo::Excelx.new(File.join(TESTDIR,'name_with_leading_slash.xlsx'))
    assert_equal 1, xlsx.sheets.count
  end
end # class
