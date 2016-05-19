require 'spec_helper'

describe Roo::ImageParser do
  describe 'parse_xlsx' do
    subject { Roo::ImageParser }
    let(:filename) { 'spec/fixtures/bar.xlsx' }
    
    context 'expect to return array of images' do
      it 'loads the proper type' do
        i_files = subject.parse_xlsx(filename)
        expect(i_files).to be_kind_of(Array)
      end
    end

    context 'expect to return empty array of images if file not exist' do
      let(:filename) { '' }

      it 'loads the proper type' do
        i_files = subject.parse_xlsx(filename)
        expect(i_files).to eql([])
      end
    end

    context 'expect to return correct images count' do
      it 'loads the proper type' do
        i_files = subject.parse_xlsx(filename)
        expect(i_files.count).to eql(19)
      end
    end
  end
end
