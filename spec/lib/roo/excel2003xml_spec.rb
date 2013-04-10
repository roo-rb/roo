require 'spec_helper'

describe Roo::Excel2003XML do
  describe '.new' do
    context 'with an xml file' do
      let(:path) { 'test/files/datetime.xml' }

      it 'loads the file' do
        expect {
          Roo::Excel2003XML.new(path)
        }.to_not raise_error
      end
    end
  end
end
