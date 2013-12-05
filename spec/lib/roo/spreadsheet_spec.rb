require 'spec_helper'

describe Roo::Spreadsheet do
  describe '.open' do
    context 'when the file extension is uppercase' do
      let(:filename) { 'file.XLS' }

      it 'loads the proper type' do
        expect(Roo::Excel).to receive(:new).with(filename, {})
        Roo::Spreadsheet.open(filename)
      end
    end

    context 'for a url' do
      context 'that is csv' do
        let(:filename) { 'http://example.com/file.csv?with=params#and=anchor' }

        it 'treats the url as CSV' do
          expect(Roo::CSV).to receive(:new).with(filename, {})
          Roo::Spreadsheet.open(filename)
        end
      end
    end

    context 'for a csv file' do
      let(:filename) { 'file.csv' }
      let(:options) { {csv_options: {col_sep: '"'}} }

      context 'with options' do
        it 'passes the options through' do
          expect(Roo::CSV).to receive(:new).with(filename, options)
          Roo::Spreadsheet.open(filename, options)
        end
      end
    end

    context 'when the file extension' do
      let(:filename) { 'file.xls' }

      context "is xls" do
        let(:options) { { extension: "xls" } }

        it 'loads with xls extension options' do
          expect(Roo::Excel).to receive(:new).with(filename, options)
          Roo::Spreadsheet.open(filename, options)
        end
      end

      context "is .xls" do
        let(:options) { { extension: ".xls" } }

        it 'loads with .xls extension options' do
          expect(Roo::Excel).to receive(:new).with(filename, options)
          Roo::Spreadsheet.open(filename, options)
        end
      end

    end
  end
end
