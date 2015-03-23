require 'spec_helper'

describe Roo::LibreOffice do
  describe '.new' do
    subject do
      Roo::LibreOffice.new('test/files/numbers1.ods')
    end

    it 'creates an instance' do
      expect(subject).to be_a(Roo::LibreOffice)
    end
  end

  describe '#sheets' do
    let(:path) { 'test/files/hidden_sheets.ods' }

    describe 'showing all sheets' do
      it 'returns the expected result' do
        expect(Roo::LibreOffice.new(path).sheets).to eq ["HiddenSheet1", "VisibleSheet1", "HiddenSheet2"]
      end
    end

    describe 'only showing visible sheets' do
      it 'returns the expected result' do
        expect(Roo::LibreOffice.new(path, only_visible_sheets: true).sheets).to eq ["VisibleSheet1"]
      end
    end
  end
end
