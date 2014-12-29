require 'spec_helper'

RSpec.describe ::Roo::Utils do
  subject { described_class }
  context '#number_to_letter' do
    ::Roo::Utils::LETTERS.each_with_index do |l, i|
      it "should return '#{l}' when passed #{i+1}" do
        expect(described_class.number_to_letter(i+1)).to eq(l)
      end
    end

    {
      27 => 'AA', 26*2 => 'AZ', 26*3 => 'BZ', 26**2 + 26 => 'ZZ', 26**2 + 27 => 'AAA',
      26**3 + 26**2 + 26 => 'ZZZ', 1.0 => 'A'
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
end
