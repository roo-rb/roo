require 'spec_helper'

RSpec.describe ::Roo::Utils do
  subject { described_class }

  context '#number_to_letter' do
    described_class::LETTERS.each_with_index do |letter, index|
      it "should return '#{ letter }' when passed #{ index + 1 }" do
        expect(described_class.number_to_letter(index + 1)).to eq(letter)
      end
    end

    {
      27 => 'AA', 26*2 => 'AZ', 26*3 => 'BZ', 26**2 + 26 => 'ZZ', 26**2 + 27 => 'AAA',
      26**3 + 26**2 + 26 => 'ZZZ', 1.0 => 'A', 676 => 'YZ', 677 => 'ZA'
    }.each do |key, value|
      it "should return '#{value}' when passed #{key}" do
        expect(described_class.number_to_letter(key)).to eq(value)
      end
    end
  end

  context '#letter_to_number' do
    it "should give 1 for 'A' and 'a'" do
      expect(described_class.letter_to_number('A')).to eq(1)
      expect(described_class.letter_to_number('a')).to eq(1)
    end

    it "should give the correct value for 'Z'" do
      expect(described_class.letter_to_number('Z')).to eq(26)
    end

    it "should give the correct value for 'AA' regardless of case mixing" do
      %w(AA aA Aa aa).each do |key|
        expect(described_class.letter_to_number(key)).to eq(27)
      end
    end

    { 'AB' => 28, 'AZ' => 26*2, 'BZ' => 26*3, 'ZZ' => 26**2 + 26 }.each do |key, value|
      it "should give the correct value for '#{key}'" do
        expect(described_class.letter_to_number(key)).to eq(value)
      end
    end
  end

  context '.split_coordinate' do
    it "returns the expected result" do
      expect(described_class.split_coordinate('A1')).to eq [1, 1]
      expect(described_class.split_coordinate('B2')).to eq [2, 2]
      expect(described_class.split_coordinate('R2')).to eq [2, 18]
      expect(described_class.split_coordinate('AR31')).to eq [31, 18 + 26]
    end
  end

  context '.extract_coordinate' do
    it "returns the expected result" do
      expect(described_class.extract_coordinate('A1')).to eq [1, 1]
      expect(described_class.extract_coordinate('B2')).to eq [2, 2]
      expect(described_class.extract_coordinate('R2')).to eq [2, 18]
      expect(described_class.extract_coordinate('AR31')).to eq [31, 18 + 26]
    end
  end

  context '.split_coord' do
    it "returns the expected result" do
      expect(described_class.split_coord('A1')).to eq ["A", 1]
      expect(described_class.split_coord('B2')).to eq ["B", 2]
      expect(described_class.split_coord('R2')).to eq ["R", 2]
      expect(described_class.split_coord('AR31')).to eq ["AR", 31]
    end

    it "raises an error when appropriate" do
      expect { described_class.split_coord('A') }.to raise_error(ArgumentError)
      expect { described_class.split_coord('2') }.to raise_error(ArgumentError)
    end
  end


  context '.num_cells_in_range' do
    it "returns the expected result" do
      expect(described_class.num_cells_in_range('A1:B2')).to eq 4
      expect(described_class.num_cells_in_range('B2:E3')).to eq 8
      expect(described_class.num_cells_in_range('R2:Z10')).to eq 81
      expect(described_class.num_cells_in_range('AR31:AR32')).to eq 2
      expect(described_class.num_cells_in_range('A1')).to eq 1
    end

    it "raises an error when appropriate" do
      expect { described_class.num_cells_in_range('A1:B1:B2') }.to raise_error(ArgumentError)
    end
  end

  context '.coordinates_in_range' do
    it "returns the expected result" do
      expect(described_class.coordinates_in_range('').to_a).to eq []
      expect(described_class.coordinates_in_range('B2').to_a).to eq [[2, 2]]
      expect(described_class.coordinates_in_range('D2:G3').to_a).to eq [[2, 4], [2, 5], [2, 6], [2, 7], [3, 4], [3, 5], [3, 6], [3, 7]]
      expect(described_class.coordinates_in_range('G3:D2').to_a).to eq []
    end

    it "raises an error when appropriate" do
      expect { described_class.coordinates_in_range('D2:G3:I5').to_a }.to raise_error(ArgumentError)
    end
  end

  context '.load_xml' do
    it 'returns the expected result' do
      expect(described_class.load_xml('test/files/sheet1.xml')).to be_a(Nokogiri::XML::Document)
      expect(described_class.load_xml('test/files/sheet1.xml').
                 remove_namespaces!.xpath("/worksheet/dimension").map do |dim|
                  dim["ref"] end.first).to eq "A1:B11"
    end
  end

  context '.each_element' do
    it 'returns the expected result' do
      described_class.each_element('test/files/sheet1.xml', 'dimension') do |dim|
        expect(dim["ref"]).to eq "A1:B11"
      end
      rows = []
      described_class.each_element('test/files/sheet1.xml', 'row') do |row|
        rows << row
      end
      expect(rows.size).to eq 11
      expect(rows[2]["r"]).to eq "3"
    end
  end
end
