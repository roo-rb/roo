require 'spec_helper'

describe Roo::CSV do
  let(:path) { 'test/files/csvtypes.csv' }
  let(:csv) { Roo::CSV.new(path) }

  describe '.new' do
    it 'creates an instance' do
      expect(csv).to be_a(Roo::CSV)
    end
  end

  describe '#parse' do
    subject do
      csv.parse(options)
    end
    context 'with headers: true' do
      let(:options) { { headers: true } }

      it "doesn't blow up" do
        expect { subject }.to_not raise_error
      end
    end
  end

  describe '#parse_with_clean_option' do
    subject do
      csv.parse(options)
    end
    context 'with clean: true' do
      let(:options) { {clean: true} }
      let(:path) { 'test/files/parse_with_clean_option.csv' }

      it "doesn't blow up" do
        expect { subject }.to_not raise_error
      end
    end
  end

  describe '#csv_options' do
    context 'when created with the csv_options option' do
      let(:options) do
        {
          col_sep: '\t',
          quote_char: "'"
        }
      end

      it 'returns the csv options' do
        csv = Roo::CSV.new(path, csv_options: options)
        expect(csv.csv_options).to eq(options)
      end
    end

    context 'when created without the csv_options option' do
      it 'returns a hash' do
        csv = Roo::CSV.new(path)
        expect(csv.csv_options).to eq({})
      end
    end
  end
end
