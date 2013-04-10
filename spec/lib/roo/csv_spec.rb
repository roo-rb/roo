require 'spec_helper'

describe Roo::Csv do
  let(:path) { 'test/files/csvtypes.csv' }

  describe '#csv_options' do
    context 'when created with the csv_options option' do
      let(:options) {
        {
          col_sep: '\t',
          quote_char: "'"
        }
      }

      it 'returns the csv options' do
        csv = Roo::Csv.new(path, csv_options: options)
        csv.csv_options.should == options
      end
    end

    context 'when created without the csv_options option' do
      it 'returns a hash' do
        csv = Roo::Csv.new(path)
        csv.csv_options.should == {}
      end
    end
  end
end
