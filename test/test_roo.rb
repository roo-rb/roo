# encoding: utf-8

# Dump warnings that come from the test to open files
# with the wrong spreadsheet class
#STDERR.reopen "/dev/null","w"

Encoding.default_external = "UTF-8"

require 'test_helper'
require 'stringio'

class TestRoo < Minitest::Test
  include TestSheets
  include TestAccesingFiles
  include TestFormulas
  include TestComments
  include TestLabels
  include TestStyles

  # Cell related tests
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

  def test_celltype
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      assert_equal :string, oo.celltype(2,6)
    end
  end

  def test_argument_error
    with_each_spreadsheet(:name=>'numbers1') do |oo|
      oo.default_sheet = "Tabelle1"
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

  # Tests for Specific Cell types (time, etc)

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

  def test_row_whitespace
    # auf dieses Dokument habe ich keinen Zugriff TODO:
    # TODO: No access to document whitespace?
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
    skip_long_test
    qq = Roo::OpenOffice.new(File.join('test', 'files', "Bibelbund.ods"))
    with_each_spreadsheet(name: 'Bibelbund') do |oo|
      oo.sheets.each do |sh|
        oo.first_row.upto(oo.last_row) do |row|
          oo.first_column.upto(oo.last_column) do |col|
            c1 = qq.cell(row, col, sh)
            c1.force_encoding("UTF-8") if c1.class == String
            c2 = oo.cell(row,col,sh)
            c2.force_encoding("UTF-8") if c2.class == String
            next if c1.nil? && c2.nil?

            assert_equal c1, c2, "diff in #{sh}/#{row}/#{col}}"
            assert_equal qq.celltype(row, col, sh), oo.celltype(row, col, sh)
            assert_equal qq.formula?(row, col, sh), oo.formula?(row, col, sh)
          end
        end
      end
    end
  end
end
