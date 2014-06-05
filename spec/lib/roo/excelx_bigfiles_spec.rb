require 'spec_helper'

describe Roo::Excelx do
  describe 'working on large files' do
    it 'reads shared strings correctly' do
      largexlsx = Roo::Excelx.new('test/files/bigfile-multitab-excelx.xlsx')

      expect(largexlsx.instance_variable_get(:@shared_table).first.should eq('Key'))
      expect(largexlsx.instance_variable_get(:@shared_table).last.should eq('HELLO-4485'))
      expect(largexlsx.instance_variable_get(:@shared_table).count.should eq(4489))
    end
    it 'reads large files quickly' do

      largexlsx = Roo::Excelx.new('test/files/bigfile-multitab-excelx.xlsx')

      largexlsx.default_sheet = 'Third'
      expect(largexlsx.cell(4197, 'A')).to eq('HELLO-4197')
      expect(largexlsx.cell(4197, 'B')).to eq(4196.02)
      expect(largexlsx.cell(4197, 'C')).to be_a(DateTime)
      expect(largexlsx.cell(4197, 'C').strftime('%Y-%m-%d %H:%M:%S')).to eq('2017-02-25 05:00:00')

      largexlsx.default_sheet = 'Second'
      expect(largexlsx.cell(2, 'A')).to eq('Google')
      expect(largexlsx.cell(2, 'A')).to be_a(Spreadsheet::Link)
      expect(largexlsx.cell(2, 'A').href).to eq('http://www.google.com')
    end
  end
end