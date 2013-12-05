# encoding: utf-8
# damit keine falschen Vermutungen aufkommen: Ich habe religioes rein gar nichts
# mit diesem Bibelbund zu tun, aber die hatten eine ziemlich grosse
# Spreadsheet-Datei mit ca. 3500 Zeilen oeffentlich im Netz, die sich ganz gut
# zum Testen eignete.
#
#--
# these test cases were developed to run under Linux OS, some commands
# (like 'diff') must be changed (or commented out ;-)) if you want to run
# the tests under another OS
#

#TODO
# Look at formulas in excel - does not work with date/time

# Dump warnings that come from the test to open files
# with the wrong spreadsheet class
#STDERR.reopen "/dev/null","w"

require File.dirname(__FILE__) + '/test_helper'

class TestRoo < Test::Unit::TestCase

  OPENOFFICE   = true 	# do OpenOffice-Spreadsheet Tests? (.ods files)
  EXCEL        = true  	# do Excel Tests? (.xls files)
  GOOGLE       = false 	# do Google-Spreadsheet Tests?
  EXCELX       = true  	# do Excelx Tests? (.xlsx files)
  LIBREOFFICE  = true  	# do LibreOffice tests? (.ods files)
  CSV          = true  	# do CSV tests? (.csv files)

  FORMATS = {
    excel: EXCEL,
    excelx: EXCELX,
    openoffice: OPENOFFICE,
    google: GOOGLE,
    libreoffice: LIBREOFFICE
  }

  ONLINE = false
  LONG_RUN = false

  def fixture_filename(name, format)
    case format
    when :excel
      "#{name}.xls"
    when :excelx
      "#{name}.xlsx"
    when :openoffice, :libreoffice
      "#{name}.ods"
    when :google
      key_of(name)
    end
  end

  # call a block of code for each spreadsheet type
  # and yield a reference to the roo object
  def with_each_spreadsheet(options)
    if options[:format]
      options[:format] = Array(options[:format])
      invalid_formats = options[:format] - FORMATS.keys
      unless invalid_formats.empty?
        raise "invalid spreadsheet types: #{invalid_formats.join(', ')}"
      end
    else
      options[:format] = FORMATS.keys
    end
    options[:format].each do |format|
      begin
        if FORMATS[format]
          yield Roo::Spreadsheet.open(File.join(TESTDIR,
            fixture_filename(options[:name], format)))
        end
      rescue => e
        raise e, "#{e.message} for #{format}", e.backtrace
      end
    end
  end

  # Using Date.strptime so check that it's using the method
  # with the value set in date_format
  def test_date
    with_each_spreadsheet(:name=>'numbers1', :format=>:google) do |oo|
      # should default to  DDMMYYYY
      assert oo.date?("21/11/1962")
      assert !oo.date?("11/21/1962")
      oo.date_format = '%m/%d/%Y'
      assert !oo.date?("21/11/1962")
      assert oo.date?("11/21/1962")
      oo.date_format = '%Y-%m-%d'
      assert(oo.date?("1962-11-21"))
      assert(!oo.date?("1962-21-11"))
    end
  end

  def test_sheets_csv
    if CSV
      oo = Roo::CSV.new(File.join(TESTDIR,'numbers1.csv'))
      assert_equal ["default"], oo.sheets
      assert_raise(RangeError) { oo.default_sheet = "no_sheet" }
      assert_raise(TypeError)  { oo.default_sheet = [1,2,3] }
      oo.sheets.each { |sh|
        oo.default_sheet = sh
        assert_equal sh, oo.default_sheet
      }
    end
  end

  def test_sheets
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      assert_equal ["Tabelle1","Name of Sheet 2","Sheet3","Sheet4","Sheet5"], oo.sheets
      assert_raise(RangeError) { oo.default_sheet = "no_sheet" }
      assert_raise(TypeError)  { oo.default_sheet = [1,2,3] }
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
      assert_raise(ArgumentError) { assert_equal "tata", oo.cell('a','f') }
      assert_raise(ArgumentError) { assert_equal "tata", oo.cell('f','a') }
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

  def test_libre_office
	  if LIBREOFFICE
      oo = Roo::LibreOffice.new(File.join(TESTDIR, "numbers1.ods"))
      oo.default_sheet = oo.sheets.first
      assert_equal 41, oo.cell('a',12)
	  end
  end

  def test_sheetname
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      oo.default_sheet = "Name of Sheet 2"
      assert_equal 'I am sheet 2', oo.cell('C',5)
      assert_raise(RangeError) { oo.default_sheet = "non existing sheet name" }
      assert_raise(RangeError) { oo.default_sheet = "non existing sheet name" }
      assert_raise(RangeError) { oo.cell('C',5,"non existing sheet name")}
      assert_raise(RangeError) { oo.celltype('C',5,"non existing sheet name")}
      assert_raise(RangeError) { oo.empty?('C',5,"non existing sheet name")}
      if oo.class == Roo::Excel
        assert_raise(NotImplementedError) { oo.formula?('C',5,"non existing sheet name")}
        assert_raise(NotImplementedError) { oo.formula('C',5,"non existing sheet name")}
      else
        assert_raise(RangeError) { oo.formula?('C',5,"non existing sheet name")}
        assert_raise(RangeError) { oo.formula('C',5,"non existing sheet name")}
        assert_raise(RangeError) { oo.set('C',5,42,"non existing sheet name")}
        assert_raise(RangeError) { oo.formulas("non existing sheet name")}
      end
      assert_raise(RangeError) { oo.to_yaml({},1,1,1,1,"non existing sheet name")}
    end
  end

  def test_argument_error
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      assert_nothing_raised(ArgumentError) {  oo.default_sheet = "Tabelle1" }
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
    with_each_spreadsheet(:name=>'simple_spreadsheet_from_italo', :format=>[:openoffice, :excel]) do |oo|
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
      assert_equal "1.0:float",oo.cell(3, 1).to_s+":"+oo.celltype(3, 1).to_s
      assert_equal "3.0:float",oo.cell(3, 2).to_s+":"+oo.celltype(3, 2).to_s
      assert_equal "1.0:float",oo.cell(3, 3).to_s+":"+oo.celltype(3, 3).to_s

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

  def test_formula_google
    with_each_spreadsheet(:name=>'formula', :format=>:google) do |oo|
      oo.default_sheet = oo.sheets.first
      assert_equal 1, oo.cell('A',1)
      assert_equal 2, oo.cell('A',2)
      assert_equal 3, oo.cell('A',3)
      assert_equal 4, oo.cell('A',4)
      assert_equal 5, oo.cell('A',5)
      assert_equal 6, oo.cell('A',6)
      # assert_equal 21, oo.cell('A',7)
      assert_equal 21.0, oo.cell('A',7) #TODO: better solution Fixnum/Float
      assert_equal :formula, oo.celltype('A',7)
      # assert_equal "=[Sheet2.A1]", oo.formula('C',7)
      # !!! different from formulas in OpenOffice
      #was: assert_equal "=sheet2!R[-6]C[-2]", oo.formula('C',7)
      # has Google changed their format of formulas/references to other sheets?
      assert_equal "=Sheet2!R[-6]C[-2]", oo.formula('C',7)
      assert_nil oo.formula('A',6)
      # assert_equal [[7, 1, "=SUM([.A1:.A6])"],
      #   [7, 2, "=SUM([.$A$1:.B6])"],
      #   [7, 3, "=[Sheet2.A1]"],
      #   [8, 2, "=SUM([.$A$1:.B7])"],
      # ], oo.formulas(oo.sheets.first)
      # different format than in openoffice spreadsheets:
      #was:
      # assert_equal [[7, 1, "=SUM(R[-6]C[0]:R[-1]C[0])"],
      #  [7, 2, "=SUM(R1C1:R[-1]C[0])"],
      #  [7, 3, "=sheet2!R[-6]C[-2]"],
      #  [8, 2, "=SUM(R1C1:R[-1]C[0])"]],
      #  oo.formulas(oo.sheets.first)
      assert_equal [[7, 1, "=SUM(R[-6]C:R[-1]C)"],
        [7, 2, "=SUM(R1C1:R[-1]C)"],
        [7, 3, "=Sheet2!R[-6]C[-2]"],
        [8, 2, "=SUM(R1C1:R[-1]C)"]],
        oo.formulas(oo.sheets.first)
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

  # Excel can only read the cell's value
  def test_formula_excel
    with_each_spreadsheet(:name=>'formula', :format=>:excel) do |oo|
      assert_equal 21, oo.cell('A',7)
      assert_equal 21, oo.cell('B',7)
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

  def test_excel_download_uri_and_zipped
    if EXCEL
      if ONLINE
        url = 'http://stiny-leonhard.de/bode-v1.xls.zip'
        excel = Roo::Excel.new(url, :zip)
        excel.default_sheet = excel.sheets.first
        assert_equal 'ist "e" im Nenner von H(s)', excel.cell('b', 5)
      end
    end
  end

  def test_openoffice_download_uri_and_zipped
    if OPENOFFICE
      if ONLINE
        url = 'http://spazioinwind.libero.it/s2/rata.ods.zip'
        sheet = Roo::OpenOffice.new(url, :zip)
        #has been changed: assert_equal 'ist "e" im Nenner von H(s)', sheet.cell('b', 5)
        assert_in_delta 0.001, 505.14, sheet.cell('c', 33).to_f
      end
    end
  end

  def test_excel_zipped
    if EXCEL
      oo = Roo::Excel.new(File.join(TESTDIR,"bode-v1.xls.zip"), :zip)
      assert oo
      assert_equal 'ist "e" im Nenner von H(s)', oo.cell('b', 5)
    end
  end

  def test_openoffice_zipped
    if OPENOFFICE
      begin
        oo = Roo::OpenOffice.new(File.join(TESTDIR,"bode-v1.ods.zip"), :zip)
        assert oo
        assert_equal 'ist "e" im Nenner von H(s)', oo.cell('b', 5)
      end
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
    #if EXCELX
    #    #Datei gibt es noch nicht
    #    oo = Roo::Excelx.new(File.join(TESTDIR,"Bibelbund1.xlsx"))
    #    oo.default_sheet = oo.sheets.first
    #    assert_equal "Tagebuch des Sekret\303\244rs.    Letzte Tagung 15./16.11.75 Schweiz", oo.cell(45,'A')
    #end
  end

  # "/tmp/xxxx" darf man unter Windows nicht verwenden, weil das nicht erkannt
  # wird.
  # Besser: Methode um temporaeres Dir. portabel zu bestimmen
  def test_huge_document_to_csv
    if LONG_RUN
      with_each_spreadsheet(:name=>'Bibelbund', :format=>[
        :openoffice,
        :excel,
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
      with_each_spreadsheet(:name=>'Bibelbund', :format=>[:openoffice,
          :excel,
          :excelx]) do |oo|
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
        assert_nothing_raised() {  oo.to_csv(File.join(tempdir,"emptysheet.csv"))  }
        assert_equal "", `cat #{File.join(tempdir,"emptysheet.csv")}`
      end
    end
  end

  def test_find_by_row_huge_document
    if LONG_RUN
      with_each_spreadsheet(:name=>'Bibelbund', :format=>[:openoffice,
          :excel,
          :excelx]) do |oo|
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

  def test_find_by_conditions
    if LONG_RUN
      with_each_spreadsheet(:name=>'Bibelbund', :format=>[:openoffice,
          :excel,
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
          :excel,
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
      assert_equal 1.75        , erg[1]['Sum'] unless oo.class == Roo::Excel
      assert_equal "Task 1"    , erg[1]['Comment']
    end
  end

  # Ruby-spreadsheet now allows us to at least give the current value
  # from a cell with a formula (no possible with parseexcel)
  def test_bug_false_borders_with_formulas
    with_each_spreadsheet(:name=>'false_encoding', :format=>:excel) do |oo|
      assert_equal 1, oo.first_row
      assert_equal 3, oo.last_row
      assert_equal 1, oo.first_column
      assert_equal 4, oo.last_column
    end
  end

  # We'ce added minimal formula support so we can now read these
  # though not sure how the spreadsheet reports older values....
  def test_fe
    with_each_spreadsheet(:name=>'false_encoding', :format=>:excel) do |oo|
      assert_equal Date.new(2007,11,1), oo.cell('a',1)
      #DOES NOT WORK IN EXCEL FILES: assert_equal true, oo.formula?('a',1)
      #DOES NOT WORK IN EXCEL FILES: assert_equal '=TODAY()', oo.formula('a',1)

      assert_equal Date.new(2008,2,9), oo.cell('B',1)
      #DOES NOT WORK IN EXCEL FILES: assert_equal true,               oo.formula?('B',1)
      #DOES NOT WORK IN EXCEL FILES: assert_equal "=A1+100",          oo.formula('B',1)

      assert_kind_of DateTime, oo.cell('C',1)
      #DOES NOT WORK IN EXCEL FILES: assert_equal true,               oo.formula?('C',1)
      #DOES NOT WORK IN EXCEL FILES: assert_equal "=C1",          oo.formula('C',1)

      assert_equal 'H1', oo.cell('A',2)
      assert_equal 'H2', oo.cell('B',2)
      assert_equal 'H3', oo.cell('C',2)
      assert_equal 'H4', oo.cell('D',2)
      assert_equal 'R1', oo.cell('A',3)
      assert_equal 'R2', oo.cell('B',3)
      assert_equal 'R3', oo.cell('C',3)
      assert_equal 'R4', oo.cell('D',3)
    end
  end

  def test_excel_does_not_support_formulas
    with_each_spreadsheet(:name=>'false_encoding', :format=>:excel) do |oo|
      assert_raise(NotImplementedError) { oo.formula('a',1) }
      assert_raise(NotImplementedError) { oo.formula?('a',1) }
      assert_raise(NotImplementedError) { oo.formulas(oo.sheets.first) }
    end
  end

  def get_extension(oo)
    case oo
    when Roo::OpenOffice
      ".ods"
    when Roo::Excel
      ".xls"
    when Roo::Excelx
      ".xlsx"
    when Roo::Google
      ""
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

  def test_bug_excel_numbers1_sheet5_last_row
    with_each_spreadsheet(:name=>'numbers1', :format=>:excel) do |oo|
      oo.default_sheet = "Tabelle1"
      assert_equal 1, oo.first_row
      assert_equal 18, oo.last_row
      assert_equal Roo::OpenOffice.letter_to_number('A'), oo.first_column
      assert_equal Roo::OpenOffice.letter_to_number('G'), oo.last_column
      oo.default_sheet = "Name of Sheet 2"
      assert_equal 5, oo.first_row
      assert_equal 14, oo.last_row
      assert_equal Roo::OpenOffice.letter_to_number('B'), oo.first_column
      assert_equal Roo::OpenOffice.letter_to_number('E'), oo.last_column
      oo.default_sheet = "Sheet3"
      assert_equal 1, oo.first_row
      assert_equal 1, oo.last_row
      assert_equal Roo::OpenOffice.letter_to_number('A'), oo.first_column
      assert_equal Roo::OpenOffice.letter_to_number('BA'), oo.last_column
      oo.default_sheet = "Sheet4"
      assert_equal 1, oo.first_row
      assert_equal 1, oo.last_row
      assert_equal Roo::OpenOffice.letter_to_number('A'), oo.first_column
      assert_equal Roo::OpenOffice.letter_to_number('E'), oo.last_column
      oo.default_sheet = "Sheet5"
      assert_equal 1, oo.first_row
      assert_equal 6, oo.last_row
      assert_equal Roo::OpenOffice.letter_to_number('A'), oo.first_column
      assert_equal Roo::OpenOffice.letter_to_number('E'), oo.last_column
    end
  end

  def test_should_raise_file_not_found_error
    if OPENOFFICE
      assert_raise(IOError) {
        Roo::OpenOffice.new(File.join('testnichtvorhanden','Bibelbund.ods'))
      }
    end
    if EXCEL
      assert_raise(IOError) {
        Roo::Excel.new(File.join('testnichtvorhanden','Bibelbund.xls'))
      }
    end
    if EXCELX
      assert_raise(IOError) {
        Roo::Excelx.new(File.join('testnichtvorhanden','Bibelbund.xlsx'))
      }
    end
    if GOOGLE
      # assert_raise(Net::HTTPServerException) {
      #   Google.new(key_of('testnichtvorhanden'+'Bibelbund.ods'))
      #   Google.new('testnichtvorhanden')
      # }
    end
  end

  def test_write_google
    # write.me: http://spreadsheets.google.com/ccc?key=ptu6bbahNZpY0N0RrxQbWdw&hl=en_GB
    with_each_spreadsheet(:name=>'write.me', :format=>:google) do |oo|
      oo.default_sheet = oo.sheets.first
      oo.set(1,1,"hello from the tests")
      assert_equal "hello from the tests", oo.cell(1,1)
      oo.set(1,1, 1.0)
      assert_equal 1.0, oo.cell(1,1)
    end
  end

  def test_bug_set_with_more_than_one_sheet_google
    # write.me: http://spreadsheets.google.com/ccc?key=ptu6bbahNZpY0N0RrxQbWdw&hl=en_GB
    with_each_spreadsheet(:name=>'write.me', :format=>:google) do |oo|
      content1 = 'AAA'
      content2 = 'BBB'
      oo.default_sheet = oo.sheets.first
      oo.set(1,1,content1)
      oo.default_sheet = oo.sheets[1]
      oo.set(1,1,content2) # in the second sheet
      oo.default_sheet = oo.sheets.first
      assert_equal content1, oo.cell(1,1)
      oo.default_sheet = oo.sheets[1]
      assert_equal content2, oo.cell(1,1)
    end
  end

  def test_set_with_sheet_argument_google
    with_each_spreadsheet(:name=>'write.me', :format=>:google) do |oo|
      random_row = rand(10)+1
      random_column = rand(10)+1
      content1 = 'ABC'
      content2 = 'DEF'
      oo.set(random_row,random_column,content1,oo.sheets.first)
      oo.set(random_row,random_column,content2,oo.sheets[1])
      assert_equal content1, oo.cell(random_row,random_column,oo.sheets.first)
      assert_equal content2, oo.cell(random_row,random_column,oo.sheets[1])
    end
  end

  def test_set_for_non_existing_sheet_google
    with_each_spreadsheet(:name=>'ptu6bbahNZpY0N0RrxQbWdw', :format=>:google) do |oo|
      assert_raise(RangeError) { oo.set(1,1,"dummy","no_sheet")   }
    end
  end

  def test_bug_bbu
    with_each_spreadsheet(:name=>'bbu', :format=>[:openoffice, :excelx, :excel]) do |oo|
      assert_nothing_raised() {
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
      }

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
      assert_nothing_raised {oo.to_xml}
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

  def test_bug_row_column_fixnum_float
    with_each_spreadsheet(:name=>'bug-row-column-fixnum-float', :format=>:excel) do |oo|
      assert_equal 42.5, oo.cell('b',2)
      assert_equal 43  , oo.cell('c',2)
      assert_equal ['hij',42.5, 43], oo.row(2)
      assert_equal ['def',42.5, 'nop'], oo.column(2)
    end
  end

  def test_file_warning_default
    if OPENOFFICE
      assert_raises(TypeError, "test/files/numbers1.xls is not an openoffice spreadsheet") {
        Roo::OpenOffice.new(File.join(TESTDIR,"numbers1.xls"))
      }
      assert_raises(TypeError) { Roo::OpenOffice.new(File.join(TESTDIR,"numbers1.xlsx")) }
    end
    if EXCEL
      assert_raises(TypeError) { Roo::Excel.new(File.join(TESTDIR,"numbers1.ods")) }
      assert_raises(TypeError) { Roo::Excel.new(File.join(TESTDIR,"numbers1.xlsx")) }
    end
    if EXCELX
      assert_raises(TypeError) { Roo::Excelx.new(File.join(TESTDIR,"numbers1.ods")) }
      assert_raises(TypeError) { Roo::Excelx.new(File.join(TESTDIR,"numbers1.xls")) }
    end
  end

  def test_file_warning_error
    if OPENOFFICE
      assert_raises(TypeError) { Roo::OpenOffice.new(File.join(TESTDIR,"numbers1.xls"),false,:error) }
      assert_raises(TypeError) { Roo::OpenOffice.new(File.join(TESTDIR,"numbers1.xlsx"),false,:error) }
    end
    if EXCEL
      assert_raises(TypeError) { Roo::Excel.new(File.join(TESTDIR,"numbers1.ods"),false,:error) }
      assert_raises(TypeError) { Roo::Excel.new(File.join(TESTDIR,"numbers1.xlsx"),false,:error) }
    end
    if EXCELX
      assert_raises(TypeError) { Roo::Excelx.new(File.join(TESTDIR,"numbers1.ods"),false,:error) }
      assert_raises(TypeError) { Roo::Excelx.new(File.join(TESTDIR,"numbers1.xls"),false,:error) }
    end
  end

  def test_file_warning_warning
    if OPENOFFICE
      assert_nothing_raised(TypeError) {
        assert_raises(Zip::ZipError) {
          Roo::OpenOffice.new(File.join(TESTDIR,"numbers1.xls"),false, :warning)
        }
      }
      assert_nothing_raised(TypeError) {
        assert_raises(Errno::ENOENT) {
          Roo::OpenOffice.new(File.join(TESTDIR,"numbers1.xlsx"),false, :warning)
        }
      }
    end
    if EXCEL
      assert_nothing_raised(TypeError) {
        assert_raises(Ole::Storage::FormatError) {
          Roo::Excel.new(File.join(TESTDIR,"numbers1.ods"),false, :warning)
        }
      }
      assert_nothing_raised(TypeError) {
        assert_raises(Ole::Storage::FormatError) {
          Roo::Excel.new(File.join(TESTDIR,"numbers1.xlsx"),false, :warning)
        }
      }
    end
    if EXCELX
      assert_nothing_raised(TypeError) {
        assert_raises(Errno::ENOENT) {
          Roo::Excelx.new(File.join(TESTDIR,"numbers1.ods"),false, :warning)
        }
      }
      assert_nothing_raised(TypeError) {
        assert_raises(Zip::ZipError) {
          Roo::Excelx.new(File.join(TESTDIR,"numbers1.xls"),false, :warning)
        }
      }
    end
  end

  def test_file_warning_ignore
    if OPENOFFICE
      # Files, die eigentlich OpenOffice-
      # Files sind, aber die falsche Endung haben.
      # Es soll ohne Fehlermeldung oder Warnung
      # oder Abbruch die Datei geoffnet werden

      # xls
      assert_nothing_raised() {
        Roo::OpenOffice.new(File.join(TESTDIR,"type_openoffice.xls"),false, :ignore)
      }
      # xlsx
      assert_nothing_raised() {
        Roo::OpenOffice.new(File.join(TESTDIR,"type_openoffice.xlsx"),false, :ignore)
      }
    end
    if EXCEL
      assert_nothing_raised() {
        Roo::Excel.new(File.join(TESTDIR,"type_excel.ods"),false, :ignore)
      }
      assert_nothing_raised() {
        Roo::Excel.new(File.join(TESTDIR,"type_excel.xlsx"),false, :ignore)
      }
    end
    if EXCELX
      assert_nothing_raised() {
        Roo::Excelx.new(File.join(TESTDIR,"type_excelx.ods"),false, :ignore)
      }
      assert_nothing_raised() {
        Roo::Excelx.new(File.join(TESTDIR,"type_excelx.xls"),false, :ignore)
      }
    end
  end

  def test_bug_last_row_excel
    with_each_spreadsheet(:name=>'time-test', :format=>:excel) do |oo|
      assert_equal 2, oo.last_row
    end
  end

  def test_bug_to_xml_with_empty_sheets
    with_each_spreadsheet(:name=>'emptysheets', :format=>[:openoffice, :excel]) do |oo|
      oo.sheets.each { |sheet|
        assert_equal nil, oo.first_row, "first_row not nil in sheet #{sheet}"
        assert_equal nil, oo.last_row, "last_row not nil in sheet #{sheet}"
        assert_equal nil, oo.first_column, "first_column not nil in sheet #{sheet}"
        assert_equal nil, oo.last_column, "last_column not nil in sheet #{sheet}"
        assert_equal nil, oo.first_row(sheet), "first_row not nil in sheet #{sheet}"
        assert_equal nil, oo.last_row(sheet), "last_row not nil in sheet #{sheet}"
        assert_equal nil, oo.first_column(sheet), "first_column not nil in sheet #{sheet}"
        assert_equal nil, oo.last_column(sheet), "last_column not nil in sheet #{sheet}"
      }
      assert_nothing_raised() { oo.to_xml }
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
      assert_kind_of DateTime, val
      assert_equal :datetime, oo.celltype('c',3)
      assert_equal DateTime.new(1961,11,21,12,17,18), val
      val = oo.cell('a',1)
      assert_kind_of Date, val
      assert_equal :date, oo.celltype('a',1)
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
    with_each_spreadsheet(:name=>'boolean', :format=>[:openoffice, :excel, :excelx]) do |oo|
      if oo.class == Roo::Excelx
        assert_equal "TRUE", oo.cell(1,1), "failure in "+oo.class.to_s
        assert_equal "FALSE", oo.cell(2,1), "failure in "+oo.class.to_s
      else
        assert_equal "true", oo.cell(1,1), "failure in "+oo.class.to_s
        assert_equal "false", oo.cell(2,1), "failure in "+oo.class.to_s
      end
    end
  end

  def test_cell_multiline
    with_each_spreadsheet(:name=>'paragraph', :format=>[:openoffice, :excel, :excelx]) do |oo|
      assert_equal "This is a test\nof a multiline\nCell", oo.cell(1,1)
      assert_equal "This is a test\n¶\nof a multiline\n\nCell", oo.cell(1,2)
      assert_equal "first p\n\nsecond p\n\nlast p", oo.cell(2,1)
    end
  end

  def test_cell_styles
    # styles only valid in excel spreadsheets?
    # TODO: what todo with other spreadsheet types
    with_each_spreadsheet(:name=>'style', :format=>[# :openoffice,
        :excel,
        # :excelx
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

  # If a cell has a date-like string but is preceeded by a '
  # to force that date to be treated like a string, we were getting an exception.
  # This test just checks for that exception to make sure it's not raised in this case
  def test_date_to_float_conversion
    with_each_spreadsheet(:name=>'datetime_floatconv', :format=>:excel) do |oo|
      assert_nothing_raised(NoMethodError) do
        oo.cell('a',1)
        oo.cell('a',2)
      end
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

  def test_ruby_spreadsheet_formula_bug
     with_each_spreadsheet(:name=>'formula_parse_error', :format=>:excel) do |oo|
       assert_equal '5026', oo.cell(2,3)
       assert_equal '5026', oo.cell(3,3)
     end
  end


  def test_excel_links
    with_each_spreadsheet(:name=>'link', :format=>:excel) do |oo|
      assert_equal 'Google', oo.cell(1,1)
      assert_equal 'http://www.google.com', oo.cell(1,1).url
    end
  end

  def test_excelx_links
    with_each_spreadsheet(:name=>'link', :format=>:excelx) do |oo|
      assert_equal 'Google', oo.cell(1,1)
      assert_equal 'http://www.google.com', oo.cell(1,1).url
    end
  end

  # Excel has two base date formats one from 1900 and the other from 1904.
  # There's a MS bug that 1900 base dates include an extra day due to erroneously
  # including 1900 as a leap yar.
  def test_base_dates_in_excel
    with_each_spreadsheet(:name=>'1900_base', :format=>:excel) do |oo|
      assert_equal Date.new(2009,06,15), oo.cell(1,1)
      #we don't want to to 'interpret' formulas  assert_equal Date.new(Time.now.year,Time.now.month,Time.now.day), oo.cell(2,1) #formula for TODAY()
      # if we test TODAY() we have also have to calculate
      # other date calculations
      #
      assert_equal :date, oo.celltype(1,1)
    end
    with_each_spreadsheet(:name=>'1904_base', :format=>:excel) do |oo|
      assert_equal Date.new(2009,06,15), oo.cell(1,1)
      # see comment above
      # assert_equal Date.new(Time.now.year,Time.now.month,Time.now.day), oo.cell(2,1) #formula for TODAY()
      assert_equal :date, oo.celltype(1,1)
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

  def test_bad_date
    with_each_spreadsheet(:name=>'prova', :format=>:excel) do |oo|
      assert_nothing_raised(ArgumentError) {
        assert_equal DateTime.new(2006,2,2,10,0,0), oo.cell('a',1)
      }
    end
  end

  def test_bad_excel_date
    with_each_spreadsheet(:name=>'bad_excel_date', :format=>:excel) do |oo|
      assert_nothing_raised(ArgumentError) {
        assert_equal DateTime.new(2006,2,2,10,0,0), oo.cell('a',1)
      }
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

      #assert_raises(ArgumentError) {
      assert_raises(NoMethodError) {
        # a42a is not a valid cell name, should raise ArgumentError
        assert_equal 9999, oo.a42a
      }
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
    with_each_spreadsheet(:name => 'matrix', :format => [:openoffice, :excel, :google]) do |oo|
      oo.default_sheet = oo.sheets.first
      assert_equal Matrix[
        [1.0, 2.0, 3.0],
        [4.0, 5.0, 6.0],
        [7.0, 8.0, 9.0] ], oo.to_matrix
    end
  end

  def test_matrix_selected_range
    with_each_spreadsheet(:name => 'matrix', :format=>[:excel,:openoffice,:google]) do |oo|
      oo.default_sheet = 'Sheet2'
      assert_equal Matrix[
        [1.0, 2.0, 3.0],
        [4.0, 5.0, 6.0],
        [7.0, 8.0, 9.0] ], oo.to_matrix(3,4,5,6)
    end
  end

  def test_matrix_all_nil
    with_each_spreadsheet(:name => 'matrix', :format=>[:excel,:openoffice,:google]) do |oo|
      oo.default_sheet = 'Sheet2'
      assert_equal Matrix[
        [nil, nil, nil],
        [nil, nil, nil],
        [nil, nil, nil] ], oo.to_matrix(10,10,12,12)
    end
  end

  def test_matrix_values_and_nil
    with_each_spreadsheet(:name => 'matrix', :format=>[:excel,:openoffice,:google]) do |oo|
      oo.default_sheet = 'Sheet3'
      assert_equal Matrix[
        [1.0, nil, 3.0],
        [4.0, 5.0, 6.0],
        [7.0, 8.0, nil] ], oo.to_matrix(1,1,3,3)
    end
  end

  def test_matrix_specifying_sheet
    with_each_spreadsheet(:name => 'matrix', :format => [:openoffice, :excel, :google]) do |oo|
      oo.default_sheet = oo.sheets.first
      assert_equal Matrix[
        [1.0, nil, 3.0],
        [4.0, 5.0, 6.0],
        [7.0, 8.0, nil] ], oo.to_matrix(nil, nil, nil, nil, 'Sheet3')
    end
  end

  # unter Windows soll es laut Bug-Reports nicht moeglich sein, eine Excel-Datei, die
  # mit Excel.new geoeffnet wurde nach dem Processing anschliessend zu loeschen.
  # Anmerkung: Das Spreadsheet-Gem erlaubt kein explizites Close von Spreadsheet-Dateien,
  # was verhindern koennte, das die Datei geloescht werden kann.
  # def test_bug_cannot_delete_opened_excel_sheet
  #   with_each_spreadsheet(:name=>'simple_spreadsheet') do |oo|
  #     'kopiere  nach temporaere Datei und versuche diese zu oeffnen und zu loeschen'
  #   end
  # end

  def test_bug_xlsx_reference_cell

    if EXCELX
=begin
    If cell A contains a string and cell B references cell A.  When reading the value of cell B, the result will be
"0.0" instead of the value of cell A.

With the attached test case, I ran the following code:
spreadsheet = Roo::Excelx.new('formula_string_error.xlsx')
spreadsheet.default_sheet = 'sheet1'
p "A: #{spreadsheet.cell(1, 1)}"
p "B: #{spreadsheet.cell(2, 1)}"

with the following results
"A: TestString"
"B: 0.0"

where the expected result is
"A: TestString"
"B: TestString"
=end
      xlsx = Roo::Excelx.new(File.join(TESTDIR, "formula_string_error.xlsx"))
      xlsx.default_sheet = xlsx.sheets.first
      assert_equal 'Teststring', xlsx.cell('a',1)
      assert_equal 'Teststring', xlsx.cell('a',2)
    end
  end

  # #formulas of an empty sheet should return an empty array and not result in
  # an error message
  # 2011-06-24
  def test_bug_formulas_empty_sheet
    with_each_spreadsheet(:name =>'emptysheets',
      :format=>[:openoffice,:excelx,:google]) do |oo|
      assert_nothing_raised(NoMethodError) {
        oo.default_sheet = oo.sheets.first
        oo.formulas
      }
      assert_equal([], oo.formulas)
    end
  end

  # #to_yaml of an empty sheet should return an empty string and not result in
  # an error message
  # 2011-06-24
  def test_bug_to_yaml_empty_sheet
    with_each_spreadsheet(:name =>'emptysheets',
      :format=>[:openoffice,:excelx,:google]) do |oo|
      assert_nothing_raised(NoMethodError) {
        oo.default_sheet = oo.sheets.first
        oo.to_yaml
      }
      assert_equal('', oo.to_yaml)
    end
  end

  # #to_matrix of an empty sheet should return an empty matrix and not result in
  # an error message
  # 2011-06-25
  def test_bug_to_matrix_empty_sheet
    with_each_spreadsheet(:name =>'emptysheets',
      :format=>[:openoffice,:excelx,:google]) do |oo|
      assert_nothing_raised(NoMethodError) {
        oo.default_sheet = oo.sheets.first
        oo.to_matrix
      }
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

  # 2011-08-11
  def test_bug_openoffice_formula_missing_letters
    if LIBREOFFICE
      # Dieses Dokument wurde mit LibreOffice angelegt.
      # Keine Ahnung, ob es damit zusammenhaengt, das diese
      # Formeln anders sind, als in der Datei formula.ods, welche
      # mit OpenOffice angelegt wurde.
      # Bei den OpenOffice-Dateien ist in diesem Feld in der XML-
      # Datei of: als Prefix enthalten, waehrend in dieser Datei
      # irgendetwas mit oooc: als Prefix verwendet wird.
      oo = Roo::OpenOffice.new(File.join(TESTDIR,'dreimalvier.ods'))
      oo.default_sheet = oo.sheets.first
      assert_equal '=SUM([.A1:.D1])', oo.formula('e',1)
      assert_equal '=SUM([.A2:.D2])', oo.formula('e',2)
      assert_equal '=SUM([.A3:.D3])', oo.formula('e',3)
      assert_equal [
       [1,5,'=SUM([.A1:.D1])'],
        [2,5,'=SUM([.A2:.D2])'],
        [3,5,'=SUM([.A3:.D3])'],
      ], oo.formulas

    end
  end

=begin
  def test_postprocessing_and_types_in_csv
    if CSV
      oo = CSV.new(File.join(TESTDIR,'csvtypes.csv'))
      oo.default_sheet = oo.sheets.first
      assert_equal(1,oo.a1)
      assert_equal(:float,oo.celltype('A',1))
      assert_equal("2",oo.b1)
      assert_equal(:string,oo.celltype('B',1))
      assert_equal("Mayer",oo.c1)
      assert_equal(:string,oo.celltype('C',1))
    end
  end
=end

=begin
  def test_postprocessing_with_callback_function
    if CSV
      oo = CSV.new(File.join(TESTDIR,'csvtypes.csv'))
      oo.default_sheet = oo.sheets.first

      #
      assert_equal(1, oo.last_column)
    end
  end
=end

=begin
  def x_123
  class ::CSV
    def cell_postprocessing(row,col,value)
      if row < 3
        return nil
      end
      return value
    end
  end
  end
=end

  def test_nil_rows_and_lines_csv
	  # x_123
	  if CSV
		  oo = Roo::CSV.new(File.join(TESTDIR,'Bibelbund.csv'))
		  oo.default_sheet = oo.sheets.first
		  assert_equal 1, oo.first_row
	  end
  end

  def test_bug_pfand_from_windows_phone_xlsx
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

  def test_comment?
    with_each_spreadsheet(:name=>'comments', :format=>[:openoffice,:libreoffice,
        :excelx]) do |oo|
      oo.default_sheet = oo.sheets.first
      assert_equal true, oo.comment?('b',4)
      assert_equal false, oo.comment?('b',99)
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
  end

  ## PREVIOUSLY SKIPPED

  # don't have these test files so removing. We can easily add in
  # by modifying with_each_spreadsheet
  GNUMERIC_ODS = false  # do gnumeric with ods files Tests?
  OPENOFFICEWRITE = false # experimental: write access with OO-Documents

  def test_writeopenoffice
    if OPENOFFICEWRITE
      File.cp(File.join(TESTDIR,"numbers1.ods"),
        File.join(TESTDIR,"numbers2.ods"))
      File.cp(File.join(TESTDIR,"numbers2.ods"),
        File.join(TESTDIR,"bak_numbers2.ods"))
      oo = OpenOffice.new(File.join(TESTDIR,"numbers2.ods"))
      oo.default_sheet = oo.sheets.first
      oo.first_row.upto(oo.last_row) {|y|
        oo.first_column.upto(oo.last_column) {|x|
          unless oo.empty?(y,x)
            # oo.set(y, x, oo.cell(y,x) + 7) if oo.celltype(y,x) == "float"
            oo.set(y, x, oo.cell(y,x) + 7) if oo.celltype(y,x) == :float
          end
        }
      }
      oo.save

      oo1 = Roo::OpenOffice.new(File.join(TESTDIR,"numbers2.ods"))
      oo2 = Roo::OpenOffice.new(File.join(TESTDIR,"bak_numbers2.ods"))
      #p oo2.to_s
      assert_equal 999, oo2.cell('a',1), oo2.cell('a',1)
      assert_equal oo2.cell('a',1) + 7, oo1.cell('a',1)
      assert_equal oo2.cell('b',1)+7, oo1.cell('b',1)
      assert_equal oo2.cell('c',1)+7, oo1.cell('c',1)
      assert_equal oo2.cell('d',1)+7, oo1.cell('d',1)
      assert_equal oo2.cell('a',2)+7, oo1.cell('a',2)
      assert_equal oo2.cell('b',2)+7, oo1.cell('b',2)
      assert_equal oo2.cell('c',2)+7, oo1.cell('c',2)
      assert_equal oo2.cell('d',2)+7, oo1.cell('d',2)
      assert_equal oo2.cell('e',2)+7, oo1.cell('e',2)

      File.cp(File.join(TESTDIR,"bak_numbers2.ods"),
        File.join(TESTDIR,"numbers2.ods"))
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

  # def test_false_encoding
  #   ex = Roo::Excel.new(File.join(TESTDIR,'false_encoding.xls'))
  #   ex.default_sheet = ex.sheets.first
  #   assert_equal "Sheet1", ex.sheets.first
  #   ex.first_row.upto(ex.last_row) do |row|
  #     ex.first_column.upto(ex.last_column) do |col|
  #       content = ex.cell(row,col)
  #       puts "#{row}/#{col}"
  #       #puts content if ! ex.empty?(row,col) or ex.formula?(row,col)
  #       if ex.formula?(row,col)
  #         #! ex.empty?(row,col)
  #         puts content
  #       end
  #     end
  #   end
  # end

  def test_simple_google
    if GOOGLE
      go = Roo::Google.new("egal")
      assert_equal "42", go.cell(1,1)
    end
  end

  def test_download_uri
    if ONLINE
      if OPENOFFICE
        assert_raises(RuntimeError) {
          Roo::OpenOffice.new("http://gibbsnichtdomainxxxxx.com/file.ods")
        }
      end
      if EXCEL
        assert_raises(RuntimeError) {
          Roo::Excel.new("http://gibbsnichtdomainxxxxx.com/file.xls")
        }
      end
      if EXCELX
        assert_raises(RuntimeError) {
          Roo::Excelx.new("http://gibbsnichtdomainxxxxx.com/file.xlsx")
        }
      end
    end
  end

  def test_download_uri_with_query_string
    dir = File.expand_path("#{File.dirname __FILE__}/files")
    { xls:  [EXCEL,       Roo::Excel],
      xlsx: [EXCELX,      Roo::Excelx],
      ods:  [OPENOFFICE,  Roo::OpenOffice]}.each do |extension, (flag, type)|
        if flag
          file = "#{dir}/simple_spreadsheet.#{extension}"
          url = "http://test.example.com/simple_spreadsheet.#{extension}?query-param=value"
          stub_request(:any, url).to_return(body: File.read(file))
          spreadsheet = type.new(url)
          spreadsheet.default_sheet = spreadsheet.sheets.first
          assert_equal 'Task 1', spreadsheet.cell('f', 4)
        end
      end
  end

  # def test_soap_server
  #   #threads = []
  #   #threads << Thread.new("serverthread") do
  #   fork do
  #     p "serverthread started"
  #     puts "in child, pid = #$$"
  #     puts `/usr/bin/ruby rooserver.rb`
  #     p "serverthread finished"
  #   end
  #   #threads << Thread.new("clientthread") do
  #   p "clientthread started"
  #   sleep 10
  #   proxy = SOAP::RPC::Driver.new("http://localhost:12321","spreadsheetserver")
  #   proxy.add_method('cell','row','col')
  #   proxy.add_method('officeversion')
  #   proxy.add_method('last_row')
  #   proxy.add_method('last_column')
  #   proxy.add_method('first_row')
  #   proxy.add_method('first_column')
  #   proxy.add_method('sheets')
  #   proxy.add_method('set_default_sheet','s')
  #   proxy.add_method('ferien_fuer_region', 'region')

  #   sheets = proxy.sheets
  #   p sheets
  #   proxy.set_default_sheet(sheets.first)

  #   assert_equal 1, proxy.first_row
  #   assert_equal 1, proxy.first_column
  #   assert_equal 187, proxy.last_row
  #   assert_equal 7, proxy.last_column
  #   assert_equal 42, proxy.cell('C',8)
  #   assert_equal 43, proxy.cell('F',12)
  #   assert_equal "1.0", proxy.officeversion
  #   p "clientthread finished"
  #   #end
  #   #threads.each {|t| t.join }
  #   puts "fertig"
  #   Process.kill("INT",pid)
  #   pid = Process.wait
  #   puts "child terminated, pid= #{pid}, status= #{$?.exitstatus}"
  # end

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

  #def sum(s,expression)
  #  arg = expression.split(':')
  #  b,z = split_coord(arg[0])
  #  first_row = z
  #  first_col = OpenOffice.letter_to_number(b)
  #  b,z = split_coord(arg[1])
  #  last_row = z
  #  last_col = OpenOffice.letter_to_number(b)
  #  result = 0
  #  first_row.upto(last_row) {|row|
  #    first_col.upto(last_col) {|col|
  #      result = result + s.cell(row,col)
  #    }
  #  }
  #  result
  #end

  #def test_dsl
  #  s = OpenOffice.new(File.join(TESTDIR,"numbers1.ods"))
  #  s.default_sheet = s.sheets.first
  #
  #    s.set 'a',1, 5
  #    s.set 'b',1, 3
  #    s.set 'c',1, 7
  #    s.set('a',2, s.cell('a',1)+s.cell('b',1))
  #    assert_equal 8, s.cell('a',2)
  #
  #    assert_equal 15, sum(s,'A1:C1')
  #  end

  #def test_create_spreadsheet1
  #  name = File.join(TESTDIR,'createdspreadsheet.ods')
  #  rm(name) if File.exists?(File.join(TESTDIR,'createdspreadsheet.ods'))
  #  # anlegen, falls noch nicht existierend
  #  s = OpenOffice.new(name,true)
  #  assert File.exists?(name)
  #end

  #def test_create_spreadsheet2
  #  # anlegen, falls noch nicht existierend
  #  s = OpenOffice.new(File.join(TESTDIR,"createdspreadsheet.ods"),true)
  #  s.set 'a',1,42
  #  s.set 'b',1,43
  #  s.set 'c',1,44
  #  s.save
  #
  #  t = OpenOffice.new(File.join(TESTDIR,"createdspreadsheet.ods"))
  #  assert_equal 42, t.cell(1,'a')
  #  assert_equal 43, t.cell('b',1)
  #  assert_equal 44, t.cell('c',3)
  #end

  # We don't have the bode-v1.xlsx test file
  # #TODO: xlsx-Datei anpassen!
  # def test_excelx_download_uri_and_zipped
  #   #TODO: gezippte xlsx Datei online zum Testen suchen
  #   if EXCELX
  #     if ONLINE
  #       url = 'http://stiny-leonhard.de/bode-v1.xlsx.zip'
  #       excel = Roo::Excelx.new(url, :zip)
  #       assert_equal 'ist "e" im Nenner von H(s)', excel.cell('b', 5)
  #     end
  #   end
  # end

  # def test_excelx_zipped
  #   # TODO: bode...xls bei Gelegenheit nach .xlsx konverieren lassen und zippen!
  #   if EXCELX
  #     # diese Datei gibt es noch nicht gezippt
  #     excel = Roo::Excelx.new(File.join(TESTDIR,"bode-v1.xlsx.zip"), :zip)
  #     assert excel
  #     assert_raises(ArgumentError) {
  #       assert_equal 'ist "e" im Nenner von H(s)', excel.cell('b', 5)
  #     }
  #     excel.default_sheet = excel.sheets.first
  #     assert_equal 'ist "e" im Nenner von H(s)', excel.cell('b', 5)
  #   end
  # end

  def test_csv_parsing_with_headers
    return unless CSV
    headers = ["TITEL", "VERFASSER", "OBJEKT", "NUMMER", "SEITE", "INTERNET", "PC", "KENNUNG"]

    oo = Roo::Spreadsheet.open(File.join(TESTDIR, 'Bibelbund.csv'))
    parsed = oo.parse(:headers => true)
    assert_equal headers, parsed[1].keys
  end

  def test_bug_numbered_sheet_names
    with_each_spreadsheet(:name=>'bug-numbered-sheet-names', :format=>:excelx) do |oo|
      assert_nothing_raised() { oo.each_with_pagename { } }
    end
  end

end # class
