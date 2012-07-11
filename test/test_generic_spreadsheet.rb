require File.dirname(__FILE__) + '/test_helper.rb'

class TestGenericSpreadsheet < Test::Unit::TestCase

  def setup
    @klass = Class.new(Roo::GenericSpreadsheet) do
      def initialize(filename='some_file')
        @cells_read = {}
        @cell = Hash.new{|h,k| h[k] = {}}
        @cell_type = Hash.new{|h,k| h[k] = {}}
        @first_row = Hash.new
        @last_row = Hash.new
        @first_column = Hash.new
        @last_column = Hash.new
        @default_sheet = 'my_sheet'
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

  def test_letters
    assert_equal 1, Roo::GenericSpreadsheet.letter_to_number('A')
    assert_equal 1, Roo::GenericSpreadsheet.letter_to_number('a')
    assert_equal 2, Roo::GenericSpreadsheet.letter_to_number('B')
    assert_equal 26, Roo::GenericSpreadsheet.letter_to_number('Z')
    assert_equal 27, Roo::GenericSpreadsheet.letter_to_number('AA')
    assert_equal 27, Roo::GenericSpreadsheet.letter_to_number('aA')
    assert_equal 27, Roo::GenericSpreadsheet.letter_to_number('Aa')
    assert_equal 27, Roo::GenericSpreadsheet.letter_to_number('aa')
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

protected
  def setup_test_sheet(workbook=nil)
    workbook ||= @oo
    %w{sheet_values sheet_types cells_read}.each do |meth|
      send("set_#{meth}".to_sym,workbook)
    end
  end

  def set_sheet_values(workbook)
    vals = workbook.instance_variable_get(:@cell)
    vals[workbook.default_sheet][[5,1]] = Date.civil(1961,11,21).to_s

    vals[workbook.default_sheet][[8,3]] = "thisisc8"
    vals[workbook.default_sheet][[8,7]] = "thisisg8"

    vals[workbook.default_sheet][[12,1]] = 41.0
    vals[workbook.default_sheet][[12,2]] = 42.0
    vals[workbook.default_sheet][[12,3]] = 43.0
    vals[workbook.default_sheet][[12,4]] = 44.0
    vals[workbook.default_sheet][[12,5]] = 45.0

    vals[workbook.default_sheet][[15,3]] = 43.0
    vals[workbook.default_sheet][[15,4]] = 44.0
    vals[workbook.default_sheet][[15,5]] = 45.0

    vals[workbook.default_sheet][[16,3]] = "dreiundvierzig"
    vals[workbook.default_sheet][[16,4]] = "vierundvierzig"
    vals[workbook.default_sheet][[16,5]] = "fuenfundvierzig"
  end

  def set_sheet_types(workbook)
    types = workbook.instance_variable_get(:@cell_type)
    types[workbook.default_sheet][[5,1]] = :date

    types[workbook.default_sheet][[8,3]] = :string
    types[workbook.default_sheet][[8,7]] = :string

    types[workbook.default_sheet][[12,1]] = :float
    types[workbook.default_sheet][[12,2]] = :float
    types[workbook.default_sheet][[12,3]] = :float
    types[workbook.default_sheet][[12,4]] = :float
    types[workbook.default_sheet][[12,5]] = :float

    types[workbook.default_sheet][[15,3]] = :float
    types[workbook.default_sheet][[15,4]] = :float
    types[workbook.default_sheet][[15,5]] = :float

    types[workbook.default_sheet][[16,3]] = :string
    types[workbook.default_sheet][[16,4]] = :string
    types[workbook.default_sheet][[16,5]] = :string
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
end