require 'spec_helper'

describe Roo::Excel do
  let(:excel) { Roo::Excel.new('test/files/boolean.xls') }

  describe '#sheets' do
    it 'returns the sheet names of the file' do
      expect(excel.sheets).to eq(["Sheet1", "Sheet2", "Sheet3"])
    end
  end
end
