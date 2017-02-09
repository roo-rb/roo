# -*- encoding: utf-8 -*-
require File.dirname(__FILE__) + '/test_helper'

class TestBase < Minitest::Test
  def setup
    @klass = Class.new(Roo::Base) do
      def initialize(filename = 'some_file')
        super
        @filename = filename
      end

      def read_cells(sheet = nil)
        @cells_read[sheet] = true
      end

      def cell(row, col, sheet = nil)
        sheet ||= default_sheet
        @cell[sheet][[row, col]]
      end

      def celltype(row, col, sheet = nil)
        sheet ||= default_sheet
        @cell_type[sheet][[row, col]]
      end

      def sheets
        ['my_sheet', 'blank sheet']
      end
    end
    @oo = @klass.new
    setup_test_sheet(@oo)
  end

  context 'private method Roo::Base.uri?(filename)' do
    should 'return true when passed a filename starts with http(s)://' do
      assert_equal true, @oo.send(:uri?, 'http://example.com/')
      assert_equal true, @oo.send(:uri?, 'https://example.com/')
    end

    should 'return false when passed a filename which does not start with http(s)://' do
      assert_equal false, @oo.send(:uri?, 'example.com')
    end

    should 'return false when passed non-String object such as Tempfile' do
      assert_equal false, @oo.send(:uri?, Tempfile.new('test'))
    end
  end

  def test_setting_invalid_type_does_not_update_cell
    @oo.set(1, 1, 1)
    assert_raises(ArgumentError) { @oo.set(1, 1, :invalid_type) }
    assert_equal 1, @oo.cell(1, 1)
    assert_equal :float, @oo.celltype(1, 1)
  end

  def test_first_row
    assert_equal 5, @oo.first_row
  end

  def test_last_row
    assert_equal 16, @oo.last_row
  end

  def test_first_column
    assert_equal 1, @oo.first_column
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

  def test_rows
    assert_equal [41.0, 42.0, 43.0, 44.0, 45.0, nil, nil], @oo.row(12)
    assert_equal [nil, '"Hello world!"', 'dreiundvierzig', 'vierundvierzig', 'fuenfundvierzig', nil, nil], @oo.row(16)
  end

  def test_empty_eh
    assert @oo.empty?(1, 1)
    assert !@oo.empty?(8, 3)
    assert @oo.empty?('A', 11)
    assert !@oo.empty?('A', 12)
  end

  def test_reload
    @oo.reload
    assert @oo.instance_variable_get(:@cell).empty?
  end

  def test_each
    oo_each = @oo.each
    assert_instance_of Enumerator, oo_each
    assert_equal [nil, '"Hello world!"', 'dreiundvierzig', 'vierundvierzig', 'fuenfundvierzig', nil, nil], oo_each.to_a.last
  end

  def test_to_yaml
    assert_equal "--- \n" + yaml_entry(5, 1, 'date', '1961-11-21'), @oo.to_yaml({}, 5, 1, 5, 1)
    assert_equal "--- \n" + yaml_entry(8, 3, 'string', 'thisisc8'), @oo.to_yaml({}, 8, 3, 8, 3)
    assert_equal "--- \n" + yaml_entry(12, 3, 'float', 43.0), @oo.to_yaml({}, 12, 3, 12, 3)
    assert_equal \
      "--- \n" + yaml_entry(12, 3, 'float', 43.0) +
        yaml_entry(12, 4, 'float', 44.0) +
        yaml_entry(12, 5, 'float', 45.0), @oo.to_yaml({}, 12, 3, 12)
    assert_equal \
      "--- \n" + yaml_entry(12, 3, 'float', 43.0) +
        yaml_entry(12, 4, 'float', 44.0) +
        yaml_entry(12, 5, 'float', 45.0) +
        yaml_entry(15, 3, 'float', 43.0) +
        yaml_entry(15, 4, 'float', 44.0) +
        yaml_entry(15, 5, 'float', 45.0) +
        yaml_entry(16, 3, 'string', 'dreiundvierzig') +
        yaml_entry(16, 4, 'string', 'vierundvierzig') +
        yaml_entry(16, 5, 'string', 'fuenfundvierzig'), @oo.to_yaml({}, 12, 3)
  end

  def test_to_csv
    assert_equal expected_csv, @oo.to_csv
  end

  def test_to_csv_with_separator
    assert_equal expected_csv_with_semicolons, @oo.to_csv(nil, ';')
  end

  protected

  def setup_test_sheet(workbook = nil)
    workbook ||= @oo
    set_sheet_values(workbook)
    set_sheet_types(workbook)
    set_cells_read(workbook)
  end

  def set_sheet_values(workbook)
    workbook.instance_variable_get(:@cell)[workbook.default_sheet] = {
      [5, 1] => Date.civil(1961, 11, 21).to_s,

      [8, 3] => 'thisisc8',
      [8, 7] => 'thisisg8',

      [12, 1] => 41.0,
      [12, 2] => 42.0,
      [12, 3] => 43.0,
      [12, 4] => 44.0,
      [12, 5] => 45.0,

      [15, 3] => 43.0,
      [15, 4] => 44.0,
      [15, 5] => 45.0,

      [16, 2] => '"Hello world!"',
      [16, 3] => 'dreiundvierzig',
      [16, 4] => 'vierundvierzig',
      [16, 5] => 'fuenfundvierzig'
    }
  end

  def set_sheet_types(workbook)
    workbook.instance_variable_get(:@cell_type)[workbook.default_sheet] = {
      [5, 1] => :date,

      [8, 3] => :string,
      [8, 7] => :string,

      [12, 1] => :float,
      [12, 2] => :float,
      [12, 3] => :float,
      [12, 4] => :float,
      [12, 5] => :float,

      [15, 3] => :float,
      [15, 4] => :float,
      [15, 5] => :float,

      [16, 2] => :string,
      [16, 3] => :string,
      [16, 4] => :string,
      [16, 5] => :string
    }
  end

  def set_first_row(workbook)
    row_hash = workbook.instance_variable_get(:@first_row)
    row_hash[workbook.default_sheet] = workbook.instance_variable_get(:@cell)[workbook.default_sheet].map { |k, _v| k[0] }.min
  end

  def set_last_row(workbook)
    row_hash = workbook.instance_variable_get(:@last_row)
    row_hash[workbook.default_sheet] = workbook.instance_variable_get(:@cell)[workbook.default_sheet].map { |k, _v| k[0] }.max
  end

  def set_first_col(workbook)
    col_hash = workbook.instance_variable_get(:@first_column)
    col_hash[workbook.default_sheet] = workbook.instance_variable_get(:@cell)[workbook.default_sheet].map { |k, _v| k[1] }.min
  end

  def set_last_col(workbook)
    col_hash = workbook.instance_variable_get(:@last_column)
    col_hash[workbook.default_sheet] = workbook.instance_variable_get(:@cell)[workbook.default_sheet].map { |k, _v| k[1] }.max
  end

  def set_cells_read(workbook)
    read_hash = workbook.instance_variable_get(:@cells_read)
    read_hash[workbook.default_sheet] = true
  end

  def expected_csv
    <<EOS
,,,,,,
,,,,,,
,,,,,,
,,,,,,
1961-11-21,,,,,,
,,,,,,
,,,,,,
,,"thisisc8",,,,"thisisg8"
,,,,,,
,,,,,,
,,,,,,
41,42,43,44,45,,
,,,,,,
,,,,,,
,,43,44,45,,
,"""Hello world!""","dreiundvierzig","vierundvierzig","fuenfundvierzig",,
EOS
  end

  def expected_csv_with_semicolons
    expected_csv.gsub(/\,/, ';')
  end
end
