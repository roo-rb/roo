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
end
