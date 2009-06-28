# These tests were all removed from test_roo.rb because they were
# from unimplemented functionality, or more commonly, missing 
# the source test data to run against. 

module SkippedTests
  # don't have these test files so removing. We can easily add in 
  # by modifying with_each_spreadsheet
  GNUMERIC_ODS = false  # do gnumeric with ods files Tests?
  OPENOFFICEWRITE = false # experimental: write access with OO-Documents

  def SKIP_test_writeopenoffice
    if OPENOFFICEWRITE
      File.cp(File.join(TESTDIR,"numbers1.ods"),
        File.join(TESTDIR,"numbers2.ods"))
      File.cp(File.join(TESTDIR,"numbers2.ods"),
        File.join(TESTDIR,"bak_numbers2.ods"))
      oo = Openoffice.new(File.join(TESTDIR,"numbers2.ods"))
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

      oo1 = Openoffice.new(File.join(TESTDIR,"numbers2.ods"))
      oo2 = Openoffice.new(File.join(TESTDIR,"bak_numbers2.ods"))
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
  
  def SKIP_test_possible_bug_snowboard_borders #no test file
    after Date.new(2008,12,15) do
      local_only do
        if EXCEL
          ex = Excel.new(File.join(TESTDIR,'problem.xls'))
          ex.default_sheet = ex.sheets.first
          assert_equal 2, ex.first_row
          assert_equal 30, ex.last_row
          assert_equal 'A', ex.first_column_as_letter
          assert_equal 'J', ex.last_column_as_letter
        end
        if EXCELX
          ex = Excelx.new(File.join(TESTDIR,'problem.xlsx'))
          ex.default_sheet = ex.sheets.first
          assert_equal 2, ex.first_row
          assert_equal 30, ex.last_row
          assert_equal 'A', ex.first_column_as_letter
          assert_equal 'J', ex.last_column_as_letter
        end
      end
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

  def SKIP_test_possible_bug_snowboard_cells # no test file
    local_only do
      after Date.new(2009,1,6) do
        # warten auf Bugfix in parseexcel
        if EXCEL
          ex = Excel.new(File.join(TESTDIR,'problem.xls'))
          ex.default_sheet = 'Custom X'
          common_possible_bug_snowboard_cells(ex)
        end
      end
      if EXCELX
        ex = Excelx.new(File.join(TESTDIR,'problem.xlsx'))
        ex.default_sheet = 'Custom X'
        common_possible_bug_snowboard_cells(ex)
      end
    end
  end

  if EXCELX
    def test_possible_bug_2008_09_13
      local_only do
        # war nur in der 1.0.0 Release ein Fehler und sollte mit aktueller
        # Release nicht mehr auftreten.
=begin
−
	<sst count="46" uniqueCount="39">
−
	0<si>
<t>Bond</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	1<si>
<t>James</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	2<si>
<t>8659</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	3<si>
<t>12B</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	4<si>
<t>087692</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	5<si>
<t>Rowe</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	6<si>
<t>Karl</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	7<si>
<t>9128</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	8<si>
<t>79A</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	9<si>
<t>Benson</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	10<si>
<t>Cedric</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	11<si>
<t>Greenstreet</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	12<si>
<t>Jenny</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	13<si>
<t>Smith</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	14<si>
<t>Greame</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	15<si>
<t>Lucas</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	16<si>
<t>Ward</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	17<si>
<t>Lee</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	18<si>
<t>Bret</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	19<si>
<t>Warne</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	20<si>
<t>Shane</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	21<si>
<t>782</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	22<si>
<t>876</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	23<si>
<t>9901</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	24<si>
<t>1235</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	25<si>
<t>16547</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	26<si>
<t>7789</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	27<si>
<t>89</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	28<si>
<t>12A</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	29<si>
<t>19A</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	30<si>
<t>256</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	31<si>
<t>129B</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	32<si>
<t>11</t>
<phoneticPr fontId="1" type="noConversion"/>
</si>
−
	33<si>
<t>Last Name</t>
</si>
−
	34<si>
<t>First Name</t>
</si>
−
35	<si>
<t>Middle Name</t>
</si>
−
	36<si>
<t>Resident ID</t>
</si>
−
	37<si>
<t>Room Number</t>
</si>
−
	38<si>
<t>Provider ID #</t>
</si>
</sst>
Hello Thomas,
How are you doing ? I am running into this strange issue with roo plugin (1.0.0). The attached
spreadsheet has all the cells formatted as "text", when I view in the Excel spreadsheet. But when it
get's into roo plugin (set_cell_values method - line 299), the values for the cells 1,1, 1,2, 1,3...1,6
show as 'date' instead of 'string'.
Because of this my parser is failing to get the proper values from the spreadsheet. Any ideas why
the formatting is getting set to the wrong value ?
Even stranger is if I save this file as ".XLS" and parse it the cells parse out fine as they are treated as
'string' instead of 'date'.
This attached file is the newer format of Microsoft Excel (.xlsx).

=end
        xx = Excelx.new(File.join(TESTDIR,'sample_file_2008-09-13.xlsx'))
        assert_equal 1, xx.sheets.size

        assert_equal 1, xx.first_row
        assert_equal 9, xx.last_row # 9 ist richtig. Es sind zwar 44 Zeilen definiert, aber der Rest hat keinen Inhalt
        assert_equal 1, xx.first_column
        assert_equal 6, xx.last_column
        assert_equal 'A', xx.first_column_as_letter
        assert_equal 'F', xx.last_column_as_letter

        assert_nothing_raised() {
          puts xx.info
        }
        p xx.cell(1,1)
        p xx.cell(1,2)
        p xx.cell(1,3)
        p xx.cell(1,4)
        p xx.cell(1,5)
        p xx.cell(1,6)
        xx.default_sheet = xx.sheets.first

        assert_equal 'Last Name', xx.cell('A',1)

        1.upto(6) do |col|
          assert_equal :string, xx.celltype(1,col)
        end
        #for col in (1..6)
        #  assert_equal "1234", xx.cell(1,col)
        #end
      end
    end
  end

  #-- bei diesen Test bekomme ich seltsamerweise einen Fehler can't allocate
  #-- memory innerhalb der zip-Routinen => erstmal deaktiviert
  def SKIP_test_huge_table_timing_10_000_openoffice #no test file
    with_each_spreadsheet(:name=>'/home/tp/ruby-test/too-testing/speedtest_10000') do |oo|    
      after Date.new(2009,1,1) do
        if LONG_RUN
          assert_nothing_raised(Timeout::Error) {
            Timeout::timeout(3.minutes) do |timeout_length|
              # process every cell
              sum = 0
              oo.sheets.each {|sheet|
                oo.default_sheet = sheet
                for row in oo.first_row..oo.last_row do
                  for col in oo.first_column..oo.last_column do
                    c = oo.cell(row,col)
                    sum += c.length if c
                  end
                end
                p sum
                assert sum > 0
              }
            end
          }
        end
      end
    end
  end

  # Eine Spreadsheetdatei wird nicht als Dateiname sondern direkt als Dokument
  # geoeffnettest_problemx_csv_imported
  def SKIP_test_from_stream_openoffice
    after Date.new(2009,1,6) do
      if OPENOFFICE
        filecontent = nil
        File.open(File.join(TESTDIR,"numbers1.ods")) do |f|
          filecontent = f.read
          p filecontent.class
          p filecontent.size
          #p filecontent
          assert filecontent.size > 0
          # #stream macht das gleiche wie #new liest abe aus Stream anstatt Datei
          oo = Openoffice.stream(filecontent)
        end
        #oo = Openoffice.open()
      end
    end
  end
  
  
  def SKIP_test_bug_encoding_exported_from_google
    if EXCEL
      xl = Excel.new(File.join(TESTDIR,"numbers1_from_google.xls"))
      xl.default_sheet = xl.sheets.first
      assert_equal 'test', xl.cell(2,'F')
    end
  end
  
  def SKIP_test_invalid_iconv_from_ms
    #TODO: does only run within a darwin-environment
    if   RUBY_PLATFORM.downcase =~ /darwin/
      assert_nothing_raised() {
        oo = Excel.new(File.join(TESTDIR,"ms.xls"))
      }
    end
  end

  def SKIP_test_false_encoding
    ex = Excel.new(File.join(TESTDIR,'false_encoding.xls'))
    ex.default_sheet = ex.sheets.first
    assert_equal "Sheet1", ex.sheets.first
    ex.first_row.upto(ex.last_row) do |row|
      ex.first_column.upto(ex.last_column) do |col|
        content = ex.cell(row,col)
        puts "#{row}/#{col}"
        #puts content if ! ex.empty?(row,col) or ex.formula?(row,col)
        if ex.formula?(row,col)
          #! ex.empty?(row,col)
          puts content
        end
      end
    end
  end

  def SKIP_test_simple_google
    if GOOGLE
      go = Google.new("egal")
      assert_equal "42", go.cell(1,1)
    end
  end
  def SKIP_test_bug_c2  # no test file
    with_each_spreadsheet(:name=>'problem', :foramt=>:excel) do |oo|
      after Date.new(2009,1,6) do
        local_only do
          expected = ['Supermodel X','T6','Shaun White','Jeremy','Custom',
            'Warhol','Twin','Malolo','Supermodel','Air','Elite',
            'King','Dominant','Dominant Slick','Blunt','Clash',
            'Bullet','Tadashi Fuse','Jussi','Royale','S-Series',
            'Fish','Love','Feelgood ES','Feelgood','GTwin','Troop',
            'Lux','Stigma','Feather','Stria','Alpha','Feelgood ICS']
          result = []
          oo.sheets[2..oo.sheets.length].each do |s|
            #(13..13).each do |s|
            oo.default_sheet = s
            name = oo.cell(2,'C')
            result << name
            #puts "#{name} (sheet: #{s})"
            #assert_equal "whatever (sheet: 13)",          "#{name} (sheet: #{s})"
          end
          assert_equal expected, result
        end
      end
    end
  end

  def SKIP_test_bug_c2_parseexcel #no test file
    after Date.new(2009,1,10) do
      local_only do
        #-- this is OK
        @workbook = Spreadsheet::ParseExcel.parse(File.join(TESTDIR,"problem.xls"))
        worksheet = @workbook.worksheet(11)
        skip = 0
        line = 1
        row = 2
        col = 3
        worksheet.each(skip) { |row_par|
          if line == row
            if row_par == nil
              raise "nil"
            end
            cell = row_par.at(col-1)
            assert cell, "cell should not be nil"
            assert_equal "Air", cell.to_s('utf-8')
          end
          line += 1
        }
        #-- worksheet 12 does not work
        @workbook = Spreadsheet::ParseExcel.parse(File.join(TESTDIR,"problem.xls"))
        worksheet = @workbook.worksheet(12)
        skip = 0
        line = 1
        row = 2
        col = 3
        worksheet.each(skip) { |row_par|
          if line == row
            if row_par == nil
              raise "nil"
            end
            cell = row_par.at(col-1)
            assert cell, "cell should not be nil"
            assert_equal "Elite", cell.to_s('utf-8')
          end
          line += 1
        }
      end
    end
  end

  def SKIP_test_bug_c2_excelx  #no test file
    after Date.new(2008,9,15) do
      local_only do
        expected = ['Supermodel X','T6','Shaun White','Jeremy','Custom',
          'Warhol','Twin','Malolo','Supermodel','Air','Elite',
          'King','Dominant','Dominant Slick','Blunt','Clash',
          'Bullet','Tadashi Fuse','Jussi','Royale','S-Series',
          'Fish','Love','Feelgood ES','Feelgood','GTwin','Troop',
          'Lux','Stigma','Feather','Stria','Alpha','Feelgood ICS']
        result = []
        @e = Excelx.new(File.join(TESTDIR,"problem.xlsx"))
        @e.sheets[2..@e.sheets.length].each do |s|
          @e.default_sheet = s
          #  assert_equal "A.",@e.cell('a',13)
          name = @e.cell(2,'C')
          result << name
          #puts "#{name} (sheet: #{s})"
          #assert_equal :string, @e.celltype('c',2)
          #assert_equal "Vapor (sheet: Vapor)", "#{name} (sheet: #{@e.sheets.first})"
          assert @e.cell(2,'c')
        end
        assert_equal expected, result

        @e = Excelx.new(File.join(TESTDIR,"problem.xlsx"))
        #@e.sheets[2..@e.sheets.length].each do |s|
        (13..13).each do |s|
          @e.default_sheet = s
          name = @e.cell(2,'C')
          #puts "#{name} (sheet: #{s})"
          assert_equal "Elite (sheet: 13)",          "#{name} (sheet: #{s})"
        end
      end
    end
  end

  def SKIP_test_compare_csv_excelx_excel  #no test file
    if EXCELX
      after Date.new(2008,12,30) do
        # parseexcel bug
        local_only do
          s1 = Excel.new(File.join(TESTDIR,"problem.xls"))
          s2 = Excelx.new(File.join(TESTDIR,"problem.xlsx"))
          s1.sheets.each {|sh| #TODO:
            s1.default_sheet = sh
            s2.default_sheet = sh
            File.delete_if_exist("/tmp/problem.csv")
            File.delete_if_exist("/tmp/problemx.csv")
            assert s1.to_csv("/tmp/problem.csv")
            assert s2.to_csv("/tmp/problemx.csv")
            assert File.exists?("/tmp/problem.csv")
            assert File.exists?("/tmp/problemx.csv")
            assert_equal "", `diff /tmp/problem.csv /tmp/problemx.csv`, "Unterschied in Sheet #{sh} #{s1.sheets.index(sh)}"
          }
        end
      end
    end
  end

  def SKIP_test_problemx_csv_imported  #no test file
    after Date.new(2009,1,6) do
      if EXCEL
        local_only do
          # wieder eingelesene CSV-Datei aus obigem Test
          # muss identisch mit problem.xls sein
          # Importieren aus csv-Datei muss manuell gemacht werden
          ex = Excel.new(File.join(TESTDIR,"problem.xls"))
          cs = Excel.new(File.join(TESTDIR,"problemx_csv_imported.xls"))
          # nur das erste sheet betrachten
          ex.default_sheet = ex.sheets.first
          cs.default_sheet = cs.sheets.first
          ex.first_row.upto(ex.last_row) do |row|
            ex.first_column.upto(ex.last_column) do |col|
              assert_equal ex.cell(row,col), cs.cell(row,col), "cell #{row}/#{col} does not match '#{ex.cell(row,col)}' '#{cs.cell(row,col)}'"
              assert_equal ex.celltype(row,col), cs.celltype(row,col), "celltype #{row}/#{col} does not match"
              assert_equal ex.empty?(row,col), cs.empty?(row,col), "empty? #{row}/#{col} does not match"
              if defined? excel_supports_formulas
                assert_equal ex.formula?(row,col), cs.formula?(row,col), "formula? #{row}/#{col} does not match"
                assert_equal ex.formula(row,col), cs.formula(row,col), "formula #{row}/#{col} does not match"
              end
            end
          end
          cs.first_row.upto(cs.last_row) do |row|
            cs.first_column.upto(cs.last_column) do |col|
              assert_equal ex.cell(row,col), cs.cell(row,col), "cell #{row}/#{col} does not match '#{ex.cell(row,col)}' '#{cs.cell(row,col)}'"
              assert_equal ex.celltype(row,col), cs.celltype(row,col), "celltype #{row}/#{col} does not match"
              assert_equal ex.empty?(row,col), cs.empty?(row,col), "empty? #{row}/#{col} does not match"
              if defined? excel_supports_formulas
                assert_equal ex.formula?(row,col), cs.formula?(row,col), "formula? #{row}/#{col} does not match"
                assert_equal ex.formula(row,col), cs.formula(row,col), "formula #{row}/#{col} does not match"
              end
            end
          end
        end
      end
    end
  end

  def SKIP_test_open_from_uri
    if ONLINE
      if OPENOFFICE
        assert_raises(RuntimeError) {
          oo = Openoffice.new("http://gibbsnichtdomainxxxxx.com/file.ods")
        }
      end
      if EXCEL
        assert_raises(RuntimeError) {
          oo = Excel.new("http://gibbsnichtdomainxxxxx.com/file.xls")
        }
      end
      if EXCELX
        assert_raises(RuntimeError) {
          oo = Excelx.new("http://gibbsnichtdomainxxxxx.com/file.xlsx")
        }
      end
    end
  end

  def SKIP_test_to_ascii_openoffice #file does not exist
    after Date.new(9999,12,31) do
      with_each_spreadsheet(:name=>'verysimple_spreadsheet', :format=>:openoffice) do |oo|    
        oo.default_sheet = oo.sheets.first
        expected="
  A    |   B   |  C   |
-------+-------+------|
      7|      8|     9|
-------+-------+------|
      4|      5|     6|
-------+-------+------|
      1|      2|     3|
----------------------/
        "
        assert_equal expected, oo.to_ascii
      end
    end
  end
  if false
    def test_soap_server
      #threads = []
      #threads << Thread.new("serverthread") do
      fork do
        p "serverthread started"
        puts "in child, pid = #$$"
        puts `/usr/bin/ruby rooserver.rb`
        p "serverthread finished"
      end
      #threads << Thread.new("clientthread") do
      p "clientthread started"
      sleep 10
      proxy = SOAP::RPC::Driver.new("http://localhost:12321","spreadsheetserver")
      proxy.add_method('cell','row','col')
      proxy.add_method('officeversion')
      proxy.add_method('last_row')
      proxy.add_method('last_column')
      proxy.add_method('first_row')
      proxy.add_method('first_column')
      proxy.add_method('sheets')
      proxy.add_method('set_default_sheet','s')
      proxy.add_method('ferien_fuer_region', 'region')

      sheets = proxy.sheets
      p sheets
      proxy.set_default_sheet(sheets.first)

      assert_equal 1, proxy.first_row
      assert_equal 1, proxy.first_column
      assert_equal 187, proxy.last_row
      assert_equal 7, proxy.last_column
      assert_equal 42, proxy.cell('C',8)
      assert_equal 43, proxy.cell('F',12)
      assert_equal "1.0", proxy.officeversion
      p "clientthread finished"
      #end
      #threads.each {|t| t.join }
      puts "fertig"
      Process.kill("INT",pid)
      pid = Process.wait
      puts "child terminated, pid= #{pid}, status= #{$?.exitstatus}"
    end
  end # false

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
  #  first_col = Openoffice.letter_to_number(b)
  #  b,z = split_coord(arg[1])
  #  last_row = z
  #  last_col = Openoffice.letter_to_number(b)
  #  result = 0
  #  first_row.upto(last_row) {|row|
  #    first_col.upto(last_col) {|col|
  #      result = result + s.cell(row,col)
  #    }
  #  }
  #  result
  #end

  #def test_dsl
  #  s = Openoffice.new(File.join(TESTDIR,"numbers1.ods"))
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
  #  s = Openoffice.new(name,true)
  #  assert File.exists?(name)
  #end

  #def test_create_spreadsheet2
  #  # anlegen, falls noch nicht existierend
  #  s = Openoffice.new(File.join(TESTDIR,"createdspreadsheet.ods"),true)
  #  s.set 'a',1,42
  #  s.set 'b',1,43
  #  s.set 'c',1,44
  #  s.save
  #
  #  #after Date.new(2007,7,3) do
  #  #  t = Openoffice.new(File.join(TESTDIR,"createdspreadsheet.ods"))
  #  #  assert_equal 42, t.cell(1,'a')
  #  #  assert_equal 43, t.cell('b',1)
  #  #  assert_equal 44, t.cell('c',3)
  #  #end
  #end
   
  #TODO: xlsx-Datei anpassen!
  def test_excelx_open_from_uri_and_zipped
    #TODO: gezippte xlsx Datei online zum Testen suchen
    after Date.new(2999,6,30) do
      if EXCELX
        if ONLINE
          url = 'http://stiny-leonhard.de/bode-v1.xlsx.zip'
          excel = Excelx.new(url, :zip)
          assert_equal 'ist "e" im Nenner von H(s)', excel.cell('b', 5)
          excel.remove_tmp # don't forget to remove the temporary files
        end
      end
    end
  end

  def test_excelx_zipped
    # TODO: bode...xls bei Gelegenheit nach .xlsx konverieren lassen und zippen!
    if EXCELX
      after Date.new(2999,7,30) do
        # diese Datei gibt es noch nicht gezippt
        excel = Excelx.new(File.join(TESTDIR,"bode-v1.xlsx.zip"), :zip)
        assert excel
        assert_raises (ArgumentError) {
          assert_equal 'ist "e" im Nenner von H(s)', excel.cell('b', 5)
        }
        excel.default_sheet = excel.sheets.first
        assert_equal 'ist "e" im Nenner von H(s)', excel.cell('b', 5)
        excel.remove_tmp # don't forget to remove the temporary files
      end
    end
  end

    
end