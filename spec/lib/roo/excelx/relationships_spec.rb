require 'spec_helper'

describe Roo::Excelx::Relationships do
  subject(:relationships) do
    Roo::Excelx::Relationships.new(path)
  end

  describe '#to_h' do
    context 'with a nil path' do
      let(:path) { nil }

      it 'returns an empty hash' do
        expect(relationships.to_h).to eq({})
      end
    end

    context 'with a non-existent file' do
      let(:path) { 'NON-EXISTANT-PATH' }

      it 'returns an empty hash' do
        expect(relationships.to_h).to eq({})
      end
    end
  end
end
