require 'spec_helper'

describe Roo::OpenOffice do
  describe '.new' do
    subject {
      Roo::OpenOffice.new('test/files/numbers1.ods')
    }

    it 'creates an instance' do
      expect(subject).to be_a(Roo::OpenOffice)
    end
  end

  # OpenOffice is an alias of LibreOffice. See libreoffice_spec.
end

describe Roo::Openoffice do
  it 'is an alias of LibreOffice' do
    expect(Roo::Openoffice).to eq(Roo::OpenOffice)
  end
end
