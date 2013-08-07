# -*- encoding: utf-8 -*-
require File.dirname(__FILE__) + '/test_helper'

class TestBase < Test::Unit::TestCase

  def setup
    @klass = Class.new(Roo::Base) do
      def initialize(filename='some_file')
        super
        @filename = filename
      end

      def read_cells(sheet=nil)
        @cells_read[sheet] = true
      end

      def cell(row, col, sheet=nil)
        sheet ||= @default_sheet
        @cell[sheet][[row,col]]
      end

      def celltype(row, col, sheet=nil)
        sheet ||= @default_sheet
        @cell_type[sheet][[row,col]]
      end

      def sheets
        ['my_sheet','blank sheet']
      end
    end
    @oo = @klass.new
    setup_test_sheet(@oo)
  end

  context 'Roo::Base.letter_to_number(letter)' do
    should "give us 1 for 'A' and 'a'" do
      assert_equal 1, Roo::Base.letter_to_number('A')
      assert_equal 1, Roo::Base.letter_to_number('a')
    end

    should "give us the correct value for 'Z'" do
      assert_equal 26, Roo::Base.letter_to_number('Z')
    end

    should "give us the correct value for 'AA' regardless of case mixing" do
      assert_equal 27, Roo::Base.letter_to_number('AA')
      assert_equal 27, Roo::Base.letter_to_number('aA')
      assert_equal 27, Roo::Base.letter_to_number('Aa')
      assert_equal 27, Roo::Base.letter_to_number('aa')
    end

    should "give us the correct value for 'AB'" do
      assert_equal 28, Roo::Base.letter_to_number('AB')
    end

    should "give us the correct value for 'AZ'" do
      assert_equal 26*2, Roo::Base.letter_to_number('AZ')
    end

    should "give us the correct value for 'BZ'" do
      assert_equal 26*3, Roo::Base.letter_to_number('BZ')
    end

    should "give us the correct value for 'ZZ'" do
      assert_equal 26**2 + 26,Roo::Base.letter_to_number('ZZ')
    end
  end

  context "Roo::Base.number_to_letter" do
    Roo::Base::LETTERS.each_with_index do |l,i|
      should "return '#{l}' when passed #{i+1}" do
        assert_equal l,Roo::Base.number_to_letter(i+1)
      end
    end

    should "return 'AA' when passed 27" do
      assert_equal 'AA',Roo::Base.number_to_letter(27)
    end

    should "return 'AZ' when passed #{26*2}" do
      assert_equal 'AZ', Roo::Base.number_to_letter(26*2)
    end

    should "return 'BZ' when passed #{26*3}" do
      assert_equal 'BZ', Roo::Base.number_to_letter(26*3)
    end

    should "return 'ZZ' when passed #{26**2 + 26}" do
      assert_equal 'ZZ',Roo::Base.number_to_letter(26**2 + 26)
    end

    should "return 'AAA' when passed #{26**2 + 27}" do
      assert_equal 'AAA',Roo::Base.number_to_letter(26**2 + 27)
    end

    should "return 'ZZZ' when passed #{26**3 + 26**2 + 26}" do
      assert_equal 'ZZZ',Roo::Base.number_to_letter(26**3 + 26**2 + 26)
    end

    should "return the correct letter when passed a Float" do
      assert_equal 'A',Roo::Base.number_to_letter(1.0)
    end
  end

  def test_setting_invalid_type_does_not_update_cell
    @oo.set(1,1,1)
    assert_raise(ArgumentError){@oo.set(1,1, :invalid_type)}
    assert_equal 1, @oo.cell(1,1)
    assert_equal :float, @oo.celltype(1,1)
  end

  def test_first_row
    assert_equal 5,@oo.first_row
  end

  def test_last_row
    assert_equal 16,@oo.last_row
  end

  def test_first_column
    assert_equal 1,@oo.first_column
  end

  def test_first_column_as_letter
    assert_equal 'A', @oo.first_column_as_letter
  end

  def test_last_column
    assert_equal 7, @oo.last_column
  end

  def test_last_column_as_letter
    assert_equal 'G', @oo.last_column_as_letter
  end

  #TODO: inkonsequente Lieferung Fixnum/Float
  def test_rows
    assert_equal [41.0,42.0,43.0,44.0,45.0, nil, nil], @oo.row(12)
    assert_equal [nil, nil, "dreiundvierzig", "vierundvierzig", "fuenfundvierzig", nil, nil], @oo.row(16)
  end

  def test_empty_eh
    assert @oo.empty?(1,1)
    assert !@oo.empty?(8,3)
    assert @oo.empty?("A",11)
    assert !@oo.empty?("A",12)
  end

  def test_reload
    @oo.reload
    assert @oo.instance_variable_get(:@cell).empty?
  end

  def test_to_yaml
    assert_equal "--- \n"+yaml_entry(5,1,"date","1961-11-21"), @oo.to_yaml({}, 5,1,5,1)
    assert_equal "--- \n"+yaml_entry(8,3,"string","thisisc8"), @oo.to_yaml({}, 8,3,8,3)
    assert_equal "--- \n"+yaml_entry(12,3,"float",43.0), @oo.to_yaml({}, 12,3,12,3)
    assert_equal \
      "--- \n"+yaml_entry(12,3,"float",43.0) +
      yaml_entry(12,4,"float",44.0) +
      yaml_entry(12,5,"float",45.0), @oo.to_yaml({}, 12,3,12)
    assert_equal \
      "--- \n"+yaml_entry(12,3,"float",43.0)+
      yaml_entry(12,4,"float",44.0)+
      yaml_entry(12,5,"float",45.0)+
      yaml_entry(15,3,"float",43.0)+
      yaml_entry(15,4,"float",44.0)+
      yaml_entry(15,5,"float",45.0)+
      yaml_entry(16,3,"string","dreiundvierzig")+
      yaml_entry(16,4,"string","vierundvierzig")+
      yaml_entry(16,5,"string","fuenfundvierzig"), @oo.to_yaml({}, 12,3)
  end

  def test_to_csv
    assert_equal expected_csv,@oo.to_csv
  end
protected
  def setup_test_sheet(workbook=nil)
    workbook ||= @oo
    set_sheet_values(workbook)
    set_sheet_types(workbook)
    set_cells_read(workbook)
  end

  def set_sheet_values(workbook)
    workbook.instance_variable_get(:@cell)[workbook.default_sheet] = {
      [5,1] => Date.civil(1961,11,21).to_s,

      [8,3] => "thisisc8",
      [8,7] => "thisisg8",

      [12,1] => 41.0,
      [12,2] => 42.0,
      [12,3] => 43.0,
      [12,4] => 44.0,
      [12,5] => 45.0,

      [15,3] => 43.0,
      [15,4] => 44.0,
      [15,5] => 45.0,

      [16,3] => "dreiundvierzig",
      [16,4] => "vierundvierzig",
      [16,5] => "fuenfundvierzig"
    }
  end

  def set_sheet_types(workbook)
    workbook.instance_variable_get(:@cell_type)[workbook.default_sheet] = {
      [5,1] => :date,

      [8,3] => :string,
      [8,7] => :string,

      [12,1] => :float,
      [12,2] => :float,
      [12,3] => :float,
      [12,4] => :float,
      [12,5] => :float,

      [15,3] => :float,
      [15,4] => :float,
      [15,5] => :float,

      [16,3] => :string,
      [16,4] => :string,
      [16,5] => :string,
    }
  end

  def set_first_row(workbook)
    row_hash = workbook.instance_variable_get(:@first_row)
    row_hash[workbook.default_sheet] = workbook.instance_variable_get(:@cell)[workbook.default_sheet].map{|k,v| k[0]}.min
  end

  def set_last_row(workbook)
    row_hash = workbook.instance_variable_get(:@last_row)
    row_hash[workbook.default_sheet] = workbook.instance_variable_get(:@cell)[workbook.default_sheet].map{|k,v| k[0]}.max
  end

  def set_first_col(workbook)
    col_hash = workbook.instance_variable_get(:@first_column)
    col_hash[workbook.default_sheet] = workbook.instance_variable_get(:@cell)[workbook.default_sheet].map{|k,v| k[1]}.min
  end

  def set_last_col(workbook)
    col_hash = workbook.instance_variable_get(:@last_column)
    col_hash[workbook.default_sheet] = workbook.instance_variable_get(:@cell)[workbook.default_sheet].map{|k,v| k[1]}.max
  end

  def set_cells_read(workbook)
    read_hash = workbook.instance_variable_get(:@cells_read)
    read_hash[workbook.default_sheet] = true
  end

  def expected_csv
    ",,,,,,\n,,,,,,\n,,,,,,\n,,,,,,\n1961-11-21,,,,,,\n,,,,,,\n,,,,,,\n,,\"thisisc8\",,,,\"thisisg8\"\n,,,,,,\n,,,,,,\n,,,,,,\n41,42,43,44,45,,\n,,,,,,\n,,,,,,\n,,43,44,45,,\n,,\"dreiundvierzig\",\"vierundvierzig\",\"fuenfundvierzig\",,\n"
  end
end
