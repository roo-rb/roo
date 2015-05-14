# encoding: utf-8
require 'spec_helper'

describe Roo::Excelx do
  subject(:xlsx) do
    Roo::Excelx.new(path)
  end

  describe '.new' do
    let(:path) { 'test/files/numeric-link.xlsx' }

    it 'creates an instance' do
      expect(subject).to be_a(Roo::Excelx)
    end

    context 'given a file with missing rels' do
      let(:path) { 'test/files/file_item_error.xlsx' }

      it 'creates an instance' do
        expect(subject).to be_a(Roo::Excelx)
      end
    end

    context 'with more cells than specified max' do
      let(:path) { 'test/files/only_one_sheet.xlsx' }

      it 'raises an appropriate error' do
        expect { Roo::Excelx.new(path, cell_max: 1) }.to raise_error(Roo::Excelx::ExceedsMaxError)
      end
    end

    context 'with fewer cells than specified max' do
      let(:path) { 'test/files/only_one_sheet.xlsx' }

      it 'creates an instance' do
        expect(Roo::Excelx.new(path, cell_max: 100)).to be_a(Roo::Excelx)
      end
    end

    context 'file path is a Pathname' do
      let(:path) { Pathname.new('test/files/file_item_error.xlsx') }

      it 'creates an instance' do
        expect(subject).to be_a(Roo::Excelx)
      end
    end

  end

  describe '#cell' do
    context 'for a link cell' do
      context 'with numeric contents' do
        let(:path) { 'test/files/numeric-link.xlsx' }

        subject do
          xlsx.cell('A', 1)
        end

        it 'returns a link with the number as a string value' do
          expect(subject).to be_a(Roo::Link)
          expect(subject).to eq('8675309.0')
        end
      end
    end

    context 'for a non-existent cell' do
      let(:path) { 'test/files/numeric-link.xlsx' }
      it 'return nil' do
        expect(xlsx.cell('AAA', 999)).to eq nil
      end
    end
  end

  describe '#parse' do
    let(:path) { 'test/files/numeric-link.xlsx' }

    context 'with a columns hash' do
      context 'when not present in the sheet' do
        it 'does not raise' do
          expect do
            xlsx.sheet(0).parse(
              this: 'This',
              that: 'That'
            )
          end.to raise_error("Couldn't find header row.")
        end
      end
    end
  end

  describe '#parse_with_clean_option' do
    let(:path) { 'test/files/parse_with_clean_option.xlsx' }
    let(:options) { {clean: true} }

    context 'with clean: true' do

      it 'does not raise' do
        expect do
          xlsx.parse(options)
        end.not_to raise_error
      end
    end
  end

  describe '#parse_unicode_with_clean_option' do
    let(:path) { 'test/files/parse_clean_with_unicode.xlsx' }
    let(:options) { {clean: true, name: 'Name'} }

    context 'with clean: true' do

      it 'returns a non empty string' do
        expect(xlsx.parse(options).last[:name]).to eql('凯')
      end
    end
  end




  describe '#sheets' do
    let(:path) { 'test/files/numbers1.xlsx' }

    it 'returns the expected result' do
      expect(subject.sheets).to eq ["Tabelle1", "Name of Sheet 2", "Sheet3", "Sheet4", "Sheet5"]
    end

    describe 'only showing visible sheets' do
      let(:path) { 'test/files/hidden_sheets.xlsx' }

      it 'returns the expected result' do
        expect(Roo::Excelx.new(path, only_visible_sheets: true).sheets).to eq ["VisibleSheet1"]
      end
    end
  end

  describe '#sheet_for' do
    let(:path) { 'test/files/numbers1.xlsx' }

    # This is kinda gross
    it 'returns the expected result' do
      expect(subject.sheet_for("Tabelle1").instance_variable_get("@name")).to eq "Tabelle1"
    end
  end

  describe '#row' do
    let(:path) { 'test/files/numbers1.xlsx' }

    it 'returns the expected result' do
      expect(subject.row(1, "Sheet5")).to eq [1.0, 5.0, 5.0, nil, nil]
    end
  end

  describe '#column' do
    let(:path) { 'test/files/numbers1.xlsx' }

    it 'returns the expected result' do
      expect(subject.column(1, "Sheet5")).to eq [1.0, 2.0, 3.0, Date.new(2007,11,21), 42.0, "ABC"]
    end
  end

  describe '#first_row' do
    let(:path) { 'test/files/numbers1.xlsx' }

    it 'returns the expected result' do
      expect(subject.first_row("Sheet5")).to eq 1
    end
  end

  describe '#last_row' do
    let(:path) { 'test/files/numbers1.xlsx' }

    it 'returns the expected result' do
      expect(subject.last_row("Sheet5")).to eq 6
    end
  end

  describe '#first_column' do
    let(:path) { 'test/files/numbers1.xlsx' }

    it 'returns the expected result' do
      expect(subject.first_column("Sheet5")).to eq 1
    end
  end

  describe '#last_column' do
    let(:path) { 'test/files/numbers1.xlsx' }

    it 'returns the expected result' do
      expect(subject.last_column("Sheet5")).to eq 5
    end
  end

  describe '#set' do

    before do
      subject.set(1, 2, "Foo", "Sheet5")
    end

    let(:path) { 'test/files/numbers1.xlsx' }
    let(:cell) { subject.cell(1, 2, "Sheet5") }

    it 'returns the expected result' do
      expect(cell).to eq "Foo"
    end
  end

  describe '#formula' do
    let(:path) { 'test/files/formula.xlsx' }

    it 'returns the expected result' do
      expect(subject.formula(1, 1, "Sheet1")).to eq nil
      expect(subject.formula(7, 2, "Sheet1")).to eq "SUM($A$1:B6)"
      expect(subject.formula(1000, 2000, "Sheet1")).to eq nil
    end
  end

  describe '#formula?' do
    let(:path) { 'test/files/formula.xlsx' }

    it 'returns the expected result' do
      expect(subject.formula?(1, 1, "Sheet1")).to eq false
      expect(subject.formula?(7, 2, "Sheet1")).to eq true
      expect(subject.formula?(1000, 2000, "Sheet1")).to eq false
    end
  end

  describe '#formulas' do
    let(:path) { 'test/files/formula.xlsx' }

    it 'returns the expected result' do
      expect(subject.formulas("Sheet1")).to eq [[7, 1, "SUM(A1:A6)"], [7, 2, "SUM($A$1:B6)"]]
    end
  end

  describe '#font' do
    let(:path) { 'test/files/style.xlsx' }

    it 'returns the expected result' do
      expect(subject.font(1, 1).bold?).to eq true
      expect(subject.font(1, 1).italic?).to eq false
      expect(subject.font(1, 1).underline?).to eq false

      expect(subject.font(7, 1).bold?).to eq false
      expect(subject.font(7, 1).italic?).to eq true
      expect(subject.font(7, 1).underline?).to eq true
      expect(subject.font(1000, 2000)).to eq nil
    end
  end

  describe '#celltype' do
    let(:path) { 'test/files/numbers1.xlsx' }

    it 'returns the expected result' do
      expect(subject.celltype(1, 1, "Sheet4")).to eq :date
      expect(subject.celltype(1, 2, "Sheet4")).to eq :float
      expect(subject.celltype(6, 2, "Sheet5")).to eq :string
      expect(subject.celltype(1000, 2000, "Sheet5")).to eq nil
    end
  end

  describe '#excelx_type' do
    let(:path) { 'test/files/numbers1.xlsx' }

    it 'returns the expected result' do
      expect(subject.excelx_type(1, 1, "Sheet5")).to eq [:numeric_or_formula, "General"]
      expect(subject.excelx_type(6, 2, "Sheet5")).to eq :string
      expect(subject.excelx_type(1000, 2000, "Sheet5")).to eq nil
    end
  end

  describe '#excelx_value' do
    let(:path) { 'test/files/numbers1.xlsx' }

    it 'returns the expected result' do
      # These values are the index in the shared strings table, might be a better
      # way to get these rather than hardcoding.
      expect(subject.excelx_value(1, 1, "Sheet5")).to eq "1"
      expect(subject.excelx_value(6, 2, "Sheet5")).to eq "16"
      expect(subject.excelx_value(6000, 2000, "Sheet5")).to eq nil
    end
  end

  describe '#excelx_format' do
    let(:path) { 'test/files/style.xlsx' }

    it 'returns the expected result' do
      # These are the index of the style for a given document
      # might be more reliable way to get this info.
      expect(subject.excelx_format(1, 1)).to eq "General"
      expect(subject.excelx_format(2, 2)).to eq "0.00"
      expect(subject.excelx_format(5000, 1000)).to eq nil
    end
  end

  describe '#empty?' do
    let(:path) { 'test/files/style.xlsx' }

    it 'returns the expected result' do
      # These are the index of the style for a given document
      # might be more reliable way to get this info.
      expect(subject.empty?(1, 1)).to eq false
      expect(subject.empty?(13, 1)).to eq true
    end
  end

  describe '#label' do
    let(:path) { 'test/files/named_cells.xlsx' }

    it 'returns the expected result' do
      expect(subject.label("berta")).to eq [4, 2, "Sheet1"]
      expect(subject.label("dave")).to eq [nil, nil, nil]
    end
  end

  describe '#labels' do
    let(:path) { 'test/files/named_cells.xlsx' }

    it 'returns the expected result' do
      expect(subject.labels).to eq [["anton", [5, 3, "Sheet1"]], ["berta", [4, 2, "Sheet1"]], ["caesar", [7, 2, "Sheet1"]]]
    end
  end

  describe '#hyperlink?' do
    let(:path) { 'test/files/link.xlsx' }

    it 'returns the expected result' do
      expect(subject.hyperlink?(1, 1)).to eq true
      expect(subject.hyperlink?(1, 2)).to eq false
    end
  end

  describe '#hyperlink' do
    let(:path) { 'test/files/link.xlsx' }

    it 'returns the expected result' do
      expect(subject.hyperlink(1, 1)).to eq "http://www.google.com"
      expect(subject.hyperlink(1, 2)).to eq nil
    end
  end

  describe '#comment' do
    let(:path) { 'test/files/comments.xlsx' }

    it 'returns the expected result' do
      expect(subject.comment(4, 2)).to eq "Kommentar fuer B4"
      expect(subject.comment(1, 2)).to eq nil
    end
  end

  describe '#comment?' do
    let(:path) { 'test/files/comments.xlsx' }

    it 'returns the expected result' do
      expect(subject.comment?(4, 2)).to eq true
      expect(subject.comment?(1, 2)).to eq false
    end
  end

  describe '#comments' do
    let(:path) { 'test/files/comments.xlsx' }

    it 'returns the expected result' do
      expect(subject.comments).to eq [[4, 2, "Kommentar fuer B4"], [5, 2, "Kommentar fuer B5"]]
    end
  end

  # nil, nil, nil, nil, nil
  # nil, nil, nil, nil, nil
  # Date	Start time	End time	Pause	Sum	Comment
  # 2007-05-07	9.25	10.25	0	1	Task 1
  # 2007-05-07	10.75	12.50	0	1.75	Task 1
  # 2007-05-07	18.00	19.00	0	1	Task 2
  # 2007-05-08	9.25	10.25	0	1	Task 2
  # 2007-05-08	14.50	15.50	0	1	Task 3
  # 2007-05-08	8.75	9.25	0	0.5	Task 3
  # 2007-05-14	21.75	22.25	0	0.5	Task 3
  # 2007-05-14	22.50	23.00	0	0.5	Task 3
  # 2007-05-15	11.75	12.75	0	1	Task 3
  # 2007-05-07	10.75	10.75	0	0	Task 1
  # nil
  describe '#each_row_streaming' do
    let(:path) { 'test/files/simple_spreadsheet.xlsx' }

    let(:expected_rows) do
      [
          [nil, nil, nil, nil, nil],
          [nil, nil, nil, nil, nil],
          ["Date", "Start time", "End time", "Pause", "Sum", "Comment", nil, nil],
          [Date.new(2007, 5, 7), 9.25, 10.25, 0.0, 1.0, "Task 1"],
          [Date.new(2007, 5, 7), 10.75, 12.50, 0.0, 1.75, "Task 1"],
          [Date.new(2007, 5, 7), 18.0, 19.0, 0.0, 1.0, "Task 2"],
          [Date.new(2007, 5, 8), 9.25, 10.25, 0.0, 1.0, "Task 2"],
          [Date.new(2007, 5, 8), 14.5, 15.5, 0.0, 1.0, "Task 3"],
          [Date.new(2007, 5, 8), 8.75, 9.25, 0.0, 0.5, "Task 3"],
          [Date.new(2007, 5, 14), 21.75, 22.25, 0.0, 0.5, "Task 3"],
          [Date.new(2007, 5, 14), 22.5, 23.0, 0.0, 0.5, "Task 3"],
          [Date.new(2007, 5, 15), 11.75, 12.75, 0.0, 1.0, "Task 3"],
          [Date.new(2007, 5, 7), 10.75, 10.75, 0.0, 0.0, "Task 1"],
          [nil]
      ]
    end

    it 'returns the expected result' do
      index = 0
      subject.each_row_streaming do |row|
        expect(row.map(&:value)).to eq expected_rows[index]
        index += 1
      end
    end

    context 'with max_rows options' do
      it 'returns the expected result' do
        index = 0
        subject.each_row_streaming(max_rows: 3) do |row|
          expect(row.map(&:value)).to eq expected_rows[index]
          index += 1
        end
        # Expect this to get incremented one time more than max (because of the increment at the end of the block)
        # but it should not be near expected_rows.size
        expect(index).to eq 4
      end
    end

    context 'with offset option' do
      let(:offset) { 3 }

      it 'returns the expected result' do
        index = 0
        subject.each_row_streaming(offset: offset) do |row|
          expect(row.map(&:value)).to eq expected_rows[index + offset]
          index += 1
        end
        expect(index).to eq 11
      end
    end

    context 'with offset and max_rows options' do
      let(:offset) { 3 }
      let(:max_rows) { 3 }

      it 'returns the expected result' do
        index = 0
        subject.each_row_streaming(offset: offset, max_rows: max_rows) do |row|
          expect(row.map(&:value)).to eq expected_rows[index + offset]
          index += 1
        end
        expect(index).to eq 4
      end
    end
  end
end
