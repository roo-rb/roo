require 'spec_helper'

describe Roo::Spreadsheet do
  describe '.open' do
    context 'when the file extension is uppercase' do
      let(:filename) { 'file.XLS' }

      it 'loads the proper type' do
        Roo::Excel.should_receive(:new).with(filename)
        Roo::Spreadsheet.open(filename)
      end
    end
  end
end
