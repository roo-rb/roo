require 'spec_helper'

describe Roo::Spreadsheet do
  describe '.open' do
    context 'when the file extension is uppercase' do
      let(:filename) { 'file.XLS' }

      it 'loads the proper type' do
        Roo::Excel.should_receive(:new).with(filename, {})
        Roo::Spreadsheet.open(filename)
      end
    end

    context 'for a csv file' do
      let(:filename) { 'file.csv' }
      let(:options) { {csv_options: {col_sep: '"'}} }

      context 'with options' do
        it 'passes the options through' do
          Roo::CSV.should_receive(:new).with(filename, options)
          Roo::Spreadsheet.open(filename, options)
        end
      end
    end
  end
end
