require 'spec_helper'

describe Roo::OpenOffice do
  describe '.new' do
    subject do
      Roo::OpenOffice.new('test/files/numbers1.ods')
    end

    it 'creates an instance' do
      expect(subject).to be_a(Roo::OpenOffice)
    end
  end

  # OpenOffice is an alias of LibreOffice. See libreoffice_spec.
end
