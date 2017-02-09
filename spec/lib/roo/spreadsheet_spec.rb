require 'spec_helper'

describe Roo::Spreadsheet do
  describe '.open' do
    context 'when the file name includes a space' do
      let(:filename) { 'great scott.xlsx' }

      it 'loads the proper type' do
        expect(Roo::Excelx).to receive(:new).with(filename, {})
        Roo::Spreadsheet.open(filename)
      end
    end

    context 'when the file extension is uppercase' do
      let(:filename) { 'file.XLSX' }

      it 'loads the proper type' do
        expect(Roo::Excelx).to receive(:new).with(filename, {})
        Roo::Spreadsheet.open(filename)
      end
    end

    context 'for a tempfile' do
      let(:tempfile) { Tempfile.new('foo.csv') }
      let(:filename) { tempfile.path }

      it 'loads the proper type' do
        expect(Roo::CSV).to receive(:new).with(filename, file_warning: :ignore).and_call_original
        expect(Roo::Spreadsheet.open(tempfile, extension: :csv)).to be_a(Roo::CSV)
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

    context 'for a windows path' do
      context 'that is xlsx' do
        let(:filename) { 'c:\Users\Joe\Desktop\myfile.xlsx' }

        it 'loads the proper type' do
          expect(Roo::Excelx).to receive(:new).with(filename, {})
          Roo::Spreadsheet.open(filename)
        end
      end
    end

    context 'for a xlsm file' do
      let(:filename) { 'macros spreadsheet.xlsm' }

      it 'loads the proper type' do
        expect(Roo::Excelx).to receive(:new).with(filename, {})
        Roo::Spreadsheet.open(filename)
      end
    end

    context 'for a csv file' do
      let(:filename) { 'file.csv' }
      let(:options) { { csv_options: { col_sep: '"' } } }

      context 'with csv_options' do
        it 'passes the csv_options through' do
          expect(Roo::CSV).to receive(:new).with(filename, options)
          Roo::Spreadsheet.open(filename, options)
        end
      end
    end

    context 'with a file extension option' do
      let(:filename) { 'file.xls' }

      context ':xlsx' do
        let(:options) { { extension: :xlsx } }

        it 'loads with xls extension options' do
          expect(Roo::Excelx).to receive(:new).with(filename, options)
          Roo::Spreadsheet.open(filename, options)
        end
      end

      context 'xlsx' do
        let(:options) { { extension: 'xlsx' } }

        it 'loads with xls extension options' do
          expect(Roo::Excelx).to receive(:new).with(filename, options)
          Roo::Spreadsheet.open(filename, options)
        end
      end

      context '.xlsx' do
        let(:options) { { extension: '.xlsx' } }

        it 'loads with .xls extension options' do
          expect(Roo::Excelx).to receive(:new).with(filename, options)
          Roo::Spreadsheet.open(filename, options)
        end
      end

    end
  end
end
