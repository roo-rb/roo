require 'spec_helper'

describe Roo::Excelx do
  describe '.new' do
    subject {
      Roo::Excelx.new('test/files/numbers1.xlsx')
    }

    it 'creates an instance' do
      expect(subject).to be_a(Roo::Excelx)
    end

    context 'given a file with missing rels' do
      subject {
        Roo::Excelx.new('test/files/file_item_error.xlsx')
      }

      it 'creates an instance' do
        expect(subject).to be_a(Roo::Excelx)
      end
    end
  end
end
