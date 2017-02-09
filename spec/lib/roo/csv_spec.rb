require 'spec_helper'

describe Roo::CSV do
  let(:path) { 'test/files/csvtypes.csv' }
  let(:csv) { Roo::CSV.new(path) }

  describe '.new' do
    it 'creates an instance' do
      expect(csv).to be_a(Roo::CSV)
    end
  end

  describe '.new with stream' do
    let(:csv) { Roo::CSV.new(File.read(path)) }
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

  describe '#set_value' do
    it 'returns the cell value' do
      expect(csv.set_value('A', 1, 'some-value', nil)).to eq('some-value')
    end
  end

  describe '#set_type' do
    it 'returns the cell type' do
      expect(csv.set_type('A', 1, 'some-type', nil)).to eq('some-type')
    end
  end
end
