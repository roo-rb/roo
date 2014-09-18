require 'spec_helper'

describe Roo::Excel do
  let(:excel) { Roo::Excel.new('test/files/boolean.xls') }

  describe '.new' do
    it 'creates an instance' do
      expect(excel).to be_a(Roo::Excel)
    end
  end

  describe '#sheets' do
    it 'returns the sheet names of the file' do
      expect(excel.sheets).to eq(["Sheet1", "Sheet2", "Sheet3"])
    end
  end

  context 'for a number cell' do
    it 'typed parse returns a Numeric' do
      expect(Roo::Excel.new('test/files/numbers1.xls').cell('A', 1)).to be_a(Numeric)
    end
    it 'untyped parse returns a String' do
      expect(Roo::Excel.new('test/files/numbers1.xls', :untyped => true).cell('A', 1)).to be_a(String)
    end
  end

end
