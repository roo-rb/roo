require 'spec_helper'

describe Roo::Base do
  let(:klass) do
    Class.new(Roo::Base) do
      def initialize(filename, data = {})
        super(filename)
        @data ||= data
      end

      def read_cells(sheet = default_sheet)
        return if @cells_read[sheet]
        type_map = { String => :string, Date => :date, Numeric => :float }

        @cell[sheet] = @data
        @cell_type[sheet] = Hash[@data.map { |k, v| [k, type_map.find {|type,_| v.is_a?(type) }.last ] }]
        @first_row[sheet] = @data.map { |k, _| k[0] }.min
        @last_row[sheet] = @data.map { |k, _| k[0] }.max
        @first_column[sheet] = @data.map { |k, _| k[1] }.min
        @last_column[sheet] = @data.map { |k, _| k[1] }.max
        @cells_read[sheet] = true
      end

      def cell(row, col, sheet = nil)
        sheet ||= default_sheet
        read_cells(sheet)
        @cell[sheet][[row, col]]
      end

      def celltype(row, col, sheet = nil)
        sheet ||= default_sheet
        read_cells(sheet)
        @cell_type[sheet][[row, col]]
      end

      def sheets
        ['my_sheet', 'blank sheet']
      end
    end
  end

  let(:spreadsheet_data) do
    {
      [3, 1] => 'Header',

      [5, 1] => Date.civil(1961, 11, 21),

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
      [16, 3] => 'forty-three',
      [16, 4] => 'forty-four',
      [16, 5] => 'forty-five'
    }
  end

  let(:spreadsheet) { klass.new('some_file', spreadsheet_data) }

  describe '#uri?' do
    it 'should return true when passed a filename starting with http(s)://' do
      expect(spreadsheet.send(:uri?, 'http://example.com/')).to be_truthy
      expect(spreadsheet.send(:uri?, 'https://example.com/')).to be_truthy
    end

    it 'should return false when passed a filename which does not start with http(s)://' do
      expect(spreadsheet.send(:uri?, 'example.com')).to be_falsy
    end

    it 'should return false when passed non-String object such as Tempfile' do
      expect(spreadsheet.send(:uri?, Tempfile.new('test'))).to be_falsy
    end
  end

  describe '#set' do
    it 'should not update cell when setting an invalid type' do
      spreadsheet.set(1, 1, 1)
      expect { spreadsheet.set(1, 1, :invalid_type) }.to raise_error(ArgumentError)
      expect(spreadsheet.cell(1, 1)).to eq(1)
      expect(spreadsheet.celltype(1, 1)).to eq(:float)
    end
  end

  describe '#first_row' do
    it 'should return the first row' do
      expect(spreadsheet.first_row).to eq(3)
    end
  end

  describe '#last_row' do
    it 'should return the last row' do
      expect(spreadsheet.last_row).to eq(16)
    end
  end

  describe '#first_column' do
    it 'should return the first column' do
      expect(spreadsheet.first_column).to eq(1)
    end
  end

  describe '#first_column_as_letter' do
    it 'should return the first column as a letter' do
      expect(spreadsheet.first_column_as_letter).to eq('A')
    end
  end

  describe '#last_column' do
    it 'should return the last column' do
      expect(spreadsheet.last_column).to eq(7)
    end
  end

  describe '#last_column_as_letter' do
    it 'should return the last column as a letter' do
      expect(spreadsheet.last_column_as_letter).to eq('G')
    end
  end

  describe "#row" do
    it "should return the specified row" do
      expect(spreadsheet.row(12)).to eq([41.0, 42.0, 43.0, 44.0, 45.0, nil, nil])
      expect(spreadsheet.row(16)).to eq([nil, '"Hello world!"', "forty-three", "forty-four", "forty-five", nil, nil])
    end

    it "should return the specified row if default_sheet is set by a string" do
      spreadsheet.default_sheet = "my_sheet"
      expect(spreadsheet.row(12)).to eq([41.0, 42.0, 43.0, 44.0, 45.0, nil, nil])
      expect(spreadsheet.row(16)).to eq([nil, '"Hello world!"', "forty-three", "forty-four", "forty-five", nil, nil])
    end

    it "should return the specified row if default_sheet is set by an integer" do
      spreadsheet.default_sheet = 0
      expect(spreadsheet.row(12)).to eq([41.0, 42.0, 43.0, 44.0, 45.0, nil, nil])
      expect(spreadsheet.row(16)).to eq([nil, '"Hello world!"', "forty-three", "forty-four", "forty-five", nil, nil])
    end
  end

  describe '#row_with' do
    context 'with a matching header row' do
      it 'returns the row number' do
        expect(spreadsheet.row_with([/Header/])). to eq 3
      end
    end

    context 'without a matching header row' do
      it 'raises an error' do
        expect { spreadsheet.row_with([/Missing Header/]) }.to \
          raise_error(Roo::HeaderRowNotFoundError)
      end

      it 'returns missing headers' do
        expect { spreadsheet.row_with([/Header/, /Missing Header 1/, /Missing Header 2/]) }.to \
          raise_error(Roo::HeaderRowNotFoundError, '[/Missing Header 1/, /Missing Header 2/]')
      end
    end
  end

  describe '#empty?' do
    it 'should return true when empty' do
      expect(spreadsheet.empty?(1, 1)).to be_truthy
      expect(spreadsheet.empty?(8, 3)).to be_falsy
      expect(spreadsheet.empty?('A', 11)).to be_truthy
      expect(spreadsheet.empty?('A', 12)).to be_falsy
    end
  end

  describe '#reload' do
    it 'should return reinitialize the spreadsheet' do
      spreadsheet.reload
      expect(spreadsheet.instance_variable_get(:@cell).empty?).to be_truthy
    end
  end

  describe '#each' do
    it 'should return an enumerator with all the rows' do
      each = spreadsheet.each
      expect(each).to be_a(Enumerator)
      expect(each.to_a.last).to eq([nil, '"Hello world!"', 'forty-three', 'forty-four', 'forty-five', nil, nil])
    end
  end

  describe "#default_sheet=" do
    it "should correctly set the default sheet if passed a string" do
      spreadsheet.default_sheet = "my_sheet"
      expect(spreadsheet.default_sheet).to eq("my_sheet")
    end

    it "should correctly set the default sheet if passed an integer" do
      spreadsheet.default_sheet = 0
      expect(spreadsheet.default_sheet).to eq("my_sheet")
    end

    it "should correctly set the default sheet if passed an integer for the second sheet" do
      spreadsheet.default_sheet = 1
      expect(spreadsheet.default_sheet).to eq("blank sheet")
    end

    it "should raise an error if passed a sheet that does not exist as an integer" do
      expect { spreadsheet.default_sheet = 10 }.to raise_error RangeError
    end

    it "should raise an error if passed a sheet that does not exist as a string" do
      expect { spreadsheet.default_sheet = "does_not_exist" }.to raise_error RangeError
    end
  end

  describe '#to_yaml' do
    it 'should convert the spreadsheet to yaml' do
      expect(spreadsheet.to_yaml({}, 5, 1, 5, 1)).to eq("--- \n" + yaml_entry(5, 1, 'date', '1961-11-21'))
      expect(spreadsheet.to_yaml({}, 8, 3, 8, 3)).to eq("--- \n" + yaml_entry(8, 3, 'string', 'thisisc8'))

      expect(spreadsheet.to_yaml({}, 12, 3, 12, 3)).to eq("--- \n" + yaml_entry(12, 3, 'float', 43.0))

      expect(spreadsheet.to_yaml({}, 12, 3, 12)).to eq(
        "--- \n" + yaml_entry(12, 3, 'float', 43.0) +
        yaml_entry(12, 4, 'float', 44.0) +
        yaml_entry(12, 5, 'float', 45.0))

      expect(spreadsheet.to_yaml({}, 12, 3)).to eq(
      "--- \n" + yaml_entry(12, 3, 'float', 43.0) +
        yaml_entry(12, 4, 'float', 44.0) +
        yaml_entry(12, 5, 'float', 45.0) +
        yaml_entry(15, 3, 'float', 43.0) +
        yaml_entry(15, 4, 'float', 44.0) +
        yaml_entry(15, 5, 'float', 45.0) +
        yaml_entry(16, 3, 'string', 'forty-three') +
        yaml_entry(16, 4, 'string', 'forty-four') +
        yaml_entry(16, 5, 'string', 'forty-five'))
    end
  end

  let(:expected_csv) do
    <<EOS
,,,,,,
,,,,,,
"Header",,,,,,
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
,"""Hello world!""","forty-three","forty-four","forty-five",,
EOS
  end

  let(:expected_csv_with_semicolons) { expected_csv.gsub(/\,/, ';') }

  describe '#to_csv' do
    it 'should convert the spreadsheet to csv' do
      expect(spreadsheet.to_csv).to eq(expected_csv)
    end

    it 'should convert the spreadsheet to csv using the separator when is passed on the parameter' do
      expect(spreadsheet.to_csv(nil, ';')).to eq(expected_csv_with_semicolons)
    end
  end
end
