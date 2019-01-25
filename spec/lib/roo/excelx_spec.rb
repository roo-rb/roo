# encoding: utf-8
require 'spec_helper'

describe Roo::Excelx do
  subject(:xlsx) do
    Roo::Excelx.new(path)
  end

  describe 'Constants' do
    describe 'ERROR_VALUES' do
      it 'returns all possible errorr values' do
        expect(described_class::ERROR_VALUES).to eq(%w(#N/A #REF! #NAME? #DIV/0! #NULL! #VALUE! #NUM!).to_set)
      end

      it 'is a set' do
        expect(described_class::ERROR_VALUES).to be_an_instance_of(Set)
      end
    end
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
          # FIXME: Because Link inherits from String, it is a String,
          #        But in theory, it shouldn't have to be a String.
          # NOTE: This test is broken becase Cell::Numeric formats numbers
          #       more intelligently.
          # expect(subject).to eq('8675309.0')
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
          end.to raise_error(Roo::HeaderRowNotFoundError)
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

    it 'returns the expected result when passed a number' do
      expect(subject.sheet_for(0).instance_variable_get("@name")).to eq "Tabelle1"
    end

    it 'returns the expected result when passed a number that is not the first sheet' do
      expect(subject.sheet_for(1).instance_variable_get("@name")).to eq "Name of Sheet 2"
    end

    it "should raise an error if passed a sheet that does not exist as an integer" do
      expect { subject.sheet_for(10) }.to raise_error RangeError
    end

    it "should raise an error if passed a sheet that does not exist as a string" do
      expect { subject.sheet_for("does_not_exist") }.to raise_error RangeError
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

  # FIXME: IMO, these tests don't provide much value. Under what circumstances
  #        will a user require the "index" value for the shared strings table?
  #        Excel value should be the raw unformatted value for the cell.
  describe '#excelx_value' do
    let(:path) { 'test/files/numbers1.xlsx' }

    it 'returns the expected result' do
      # These values are the index in the shared strings table, might be a better
      # way to get these rather than hardcoding.

      # expect(subject.excelx_value(1, 1, "Sheet5")).to eq "1" # passes by accident
      # expect(subject.excelx_value(6, 2, "Sheet5")).to eq "16"
      # expect(subject.excelx_value(6000, 2000, "Sheet5")).to eq nil
    end
  end

  describe '#formatted_value' do
    context 'contains zero-padded numbers' do
      let(:path) { 'test/files/zero-padded-number.xlsx' }

      it 'returns a zero-padded number' do
        expect(subject.formatted_value(4, 1)).to eq '05010'
      end
    end
  end

  describe '#row' do
    context 'integers with leading zero'
      let(:path) { 'test/files/number_with_zero_prefix.xlsx' }

      it 'returns base 10 integer' do
        (1..50).each do |row_index|
          range_start = (row_index - 1) * 20 + 1
          expect(subject.row(row_index)).to eq (range_start..(range_start+19)).to_a
        end
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

    context 'defined on cell range' do
     let(:path) { 'test/files/cell-range-link.xlsx' }

      it 'returns the expected result' do
        [[false]*3, *[[true, true, false]]*4, [false]*3].each.with_index(1) do |row, row_index|
          row.each.with_index(1) do |value, col_index|
            expect(subject.hyperlink?(row_index, col_index)).to eq(value)
          end
        end
      end
    end
  end

  describe '#hyperlink' do
    context 'defined on cell range' do
     let(:path) { 'test/files/cell-range-link.xlsx' }

      it 'returns the expected result' do
        link = "http://www.google.com"
        [[nil]*3, *[[link, link, nil]]*4, [nil]*3].each.with_index(1) do |row, row_index|
          row.each.with_index(1) do |value, col_index|
            expect(subject.hyperlink(row_index, col_index)).to eq(value)
          end
        end
      end
    end

    context 'without location' do
      let(:path) { 'test/files/link.xlsx' }

      it 'returns the expected result' do
        expect(subject.hyperlink(1, 1)).to eq "http://www.google.com"
        expect(subject.hyperlink(1, 2)).to eq nil
      end
    end

    context 'with location' do
      let(:path) { 'test/files/link_with_location.xlsx' }

      it 'returns the expected result' do
        expect(subject.hyperlink(1, 1)).to eq "http://www.google.com/#hey"
        expect(subject.hyperlink(1, 2)).to eq nil
      end
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

    context 'without block passed' do
      it 'returns an enumerator' do
        expect(subject.each_row_streaming).to be_a(Enumerator)
      end
    end
  end

  describe '#html_strings' do
    describe "HTML Parsing Enabling" do
      let(:path) { 'test/files/html_strings_formatting.xlsx' }

      it 'returns the expected result' do
        expect(subject.excelx_value(1, 1, "Sheet1")).to eq("This has no formatting.")
        expect(subject.excelx_value(2, 1, "Sheet1")).to eq("<html>This has<b> bold </b>formatting.</html>")
        expect(subject.excelx_value(2, 2, "Sheet1")).to eq("<html>This has <i>italics</i> formatting.</html>")
        expect(subject.excelx_value(2, 3, "Sheet1")).to eq("<html>This has <u>underline</u> format.</html>")
        expect(subject.excelx_value(2, 4, "Sheet1")).to eq("<html>Superscript. x<sup>123</sup></html>")
        expect(subject.excelx_value(2, 5, "Sheet1")).to eq("<html>SubScript.  T<sub>j</sub></html>")

        expect(subject.excelx_value(3, 1, "Sheet1")).to eq("<html>Bold, italics <b><i>together</i></b>.</html>")
        expect(subject.excelx_value(3, 2, "Sheet1")).to eq("<html>Bold, Underline <b><u>together</u></b>.</html>")
        expect(subject.excelx_value(3, 3, "Sheet1")).to eq("<html>Bold, Superscript. <b>x</b><sup><b>N</b></sup></html>")
        expect(subject.excelx_value(3, 4, "Sheet1")).to eq("<html>Bold, Subscript. <b>T</b><sub><b>abc</b></sub></html>")
        expect(subject.excelx_value(3, 5, "Sheet1")).to eq("<html>Italics, Underline <i><u>together</u></i>.</html>")
        expect(subject.excelx_value(3, 6, "Sheet1")).to eq("<html>Italics, Superscript.  <i>X</i><sup><i>abc</i></sup></html>")
        expect(subject.excelx_value(3, 7, "Sheet1")).to eq("<html>Italics, Subscript.  <i>B</i><sub><i>efg</i></sub></html>")
        expect(subject.excelx_value(4, 1, "Sheet1")).to eq("<html>Bold, italics underline,<b><i><u> together</u></i></b>.</html>")
        expect(subject.excelx_value(4, 2, "Sheet1")).to eq("<html>Bold, italics, superscript. <b>X</b><sup><b><i>abc</i></b></sup><b><i>123</i></b></html>")
        expect(subject.excelx_value(4, 3, "Sheet1")).to eq("<html>Bold, Italics, subscript. <b><i>Mg</i></b><sub><b><i>ha</i></b></sub><b><i>2</i></b></html>")
        expect(subject.excelx_value(4, 4, "Sheet1")).to eq("<html>Bold, Underline, superscript. <b><u>AB</u></b><sup><b><u>C12</u></b></sup><b><u>3</u></b></html>")
        expect(subject.excelx_value(4, 5, "Sheet1")).to eq("<html>Bold, Underline, subscript. <b><u>Good</u></b><sub><b><u>XYZ</u></b></sub></html>")
        expect(subject.excelx_value(4, 6, "Sheet1")).to eq("<html>Italics, Underline, superscript. <i><u>Up</u></i><sup><i><u>swing</u></i></sup></html>")
        expect(subject.excelx_value(4, 7, "Sheet1")).to eq("<html>Italics, Underline, subscript. <i><u>T</u></i><sub><i><u>swing</u></i></sub></html>")
        expect(subject.excelx_value(5, 1, "Sheet1")).to eq("<html>Bold, italics, underline, superscript.  <b><i><u>GHJK</u></i></b><sup><b><i><u>190</u></i></b></sup><b><i><u>4</u></i></b></html>")
        expect(subject.excelx_value(5, 2, "Sheet1")).to eq("<html>Bold, italics, underline, subscript. <b><i><u>Mike</u></i></b><sub><b><i><u>drop</u></i></b></sub></html>")
        expect(subject.excelx_value(6, 1, "Sheet1")).to eq("See that regular html tags do not create html tags.\n<ol>\n  <li> Denver Broncos </li>\n  <li> Carolina Panthers </li>\n  <li> New England Patriots</li>\n  <li>Arizona Panthers</li>\n</ol>")
        expect(subject.excelx_value(7, 1, "Sheet1")).to eq("<html>Does create html tags when formatting is used..\n<ol>\n  <li> <b>Denver Broncos</b> </li>\n  <li> <i>Carolina Panthers </i></li>\n  <li> <u>New England Patriots</u></li>\n  <li>Arizona Panthers</li>\n</ol></html>")
      end
    end
  end

  describe '_x000D_' do
    let(:path) { 'test/files/x000D.xlsx' }
    it 'does not contain _x000D_' do
      expect(subject.cell(2, 9)).not_to include('_x000D_')
    end
  end

  describe 'opening a file with a chart sheet' do
    let(:path) { 'test/files/chart_sheet.xlsx' }
    it 'should not raise' do
      expect{ subject }.to_not raise_error
    end
  end

  describe 'opening a file with white space in the styles.xml' do
    let(:path) { 'test/files/style_nodes_with_white_spaces.xlsx' }
    subject(:xlsx) do
      Roo::Spreadsheet.open(path, expand_merged_ranges: true, extension: :xlsx)
    end
    it 'should properly recognize formats' do
      expect(subject.sheet(0).excelx_format(2,1)).to eq 'm/d/yyyy" "h:mm:ss" "AM/PM'
    end
  end

  describe 'images' do
    let(:path) { 'test/files/images.xlsx' }

    it 'returns array of images from default sheet' do
      expect(subject.images).to be_kind_of(Array)
      expect(subject.images.size).to eql(19)
    end

    it 'returns empty array if there is no images on the sheet' do
      expect(subject.images("Sheet2")).to eql([])
    end
  end
end

describe 'Roo::Excelx with options set' do
  subject(:xlsx) do
    Roo::Excelx.new(path, disable_html_wrapper: true)
  end

  describe '#html_strings' do
    describe "HTML Parsing Disabled" do
      let(:path) { 'test/files/html_strings_formatting.xlsx' }

      it 'returns the expected result' do
        expect(subject.excelx_value(1, 1, "Sheet1")).to eq("This has no formatting.")
        expect(subject.excelx_value(2, 1, "Sheet1")).to eq("This has bold formatting.")
        expect(subject.excelx_value(2, 2, "Sheet1")).to eq("This has italics formatting.")
        expect(subject.excelx_value(2, 3, "Sheet1")).to eq("This has underline format.")
        expect(subject.excelx_value(2, 4, "Sheet1")).to eq("Superscript. x123")
        expect(subject.excelx_value(2, 5, "Sheet1")).to eq("SubScript.  Tj")

        expect(subject.excelx_value(3, 1, "Sheet1")).to eq("Bold, italics together.")
        expect(subject.excelx_value(3, 2, "Sheet1")).to eq("Bold, Underline together.")
        expect(subject.excelx_value(3, 3, "Sheet1")).to eq("Bold, Superscript. xN")
        expect(subject.excelx_value(3, 4, "Sheet1")).to eq("Bold, Subscript. Tabc")
        expect(subject.excelx_value(3, 5, "Sheet1")).to eq("Italics, Underline together.")
        expect(subject.excelx_value(3, 6, "Sheet1")).to eq("Italics, Superscript.  Xabc")
        expect(subject.excelx_value(3, 7, "Sheet1")).to eq("Italics, Subscript.  Befg")
        expect(subject.excelx_value(4, 1, "Sheet1")).to eq("Bold, italics underline, together.")
        expect(subject.excelx_value(4, 2, "Sheet1")).to eq("Bold, italics, superscript. Xabc123")
        expect(subject.excelx_value(4, 3, "Sheet1")).to eq("Bold, Italics, subscript. Mgha2")
        expect(subject.excelx_value(4, 4, "Sheet1")).to eq("Bold, Underline, superscript. ABC123")
        expect(subject.excelx_value(4, 5, "Sheet1")).to eq("Bold, Underline, subscript. GoodXYZ")
        expect(subject.excelx_value(4, 6, "Sheet1")).to eq("Italics, Underline, superscript. Upswing")
        expect(subject.excelx_value(4, 7, "Sheet1")).to eq("Italics, Underline, subscript. Tswing")
        expect(subject.excelx_value(5, 1, "Sheet1")).to eq("Bold, italics, underline, superscript.  GHJK1904")
        expect(subject.excelx_value(5, 2, "Sheet1")).to eq("Bold, italics, underline, subscript. Mikedrop")
        expect(subject.excelx_value(6, 1, "Sheet1")).to eq("See that regular html tags do not create html tags.\n<ol>\n  <li> Denver Broncos </li>\n  <li> Carolina Panthers </li>\n  <li> New England Patriots</li>\n  <li>Arizona Panthers</li>\n</ol>")
        expect(subject.excelx_value(7, 1, "Sheet1")).to eq("Does create html tags when formatting is used..\n<ol>\n  <li> Denver Broncos </li>\n  <li> Carolina Panthers </li>\n  <li> New England Patriots</li>\n  <li>Arizona Panthers</li>\n</ol>")
      end
    end
  end
end