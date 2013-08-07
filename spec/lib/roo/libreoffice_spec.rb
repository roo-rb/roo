require 'spec_helper'

describe Roo::Libreoffice do
  describe '.new' do
    subject {
      Roo::Libreoffice.new('test/files/numbers1.ods')
    }

    it 'creates an instance' do
      expect(subject).to be_a(Roo::Libreoffice)
    end
  end
end
