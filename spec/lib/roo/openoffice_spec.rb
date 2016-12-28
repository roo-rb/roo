require 'spec_helper'

describe Roo::OpenOffice do
  describe '.new' do
    subject do
      Roo::OpenOffice.new('test/files/numbers1.ods')
    end

    it 'creates an instance' do
      expect(subject).to be_a(Roo::OpenOffice)
    end

    context 'for float/integer values' do
      context 'integer without point' do
        it { expect(subject.cell(3,"A","Sheet4")).to eq(1234) }
        it { expect(subject.cell(3,"A","Sheet4")).to be_a(Integer) }
      end

      context 'float with point' do
        it { expect(subject.cell(3,"B","Sheet4")).to eq(1234.00) }
        it { expect(subject.cell(3,"B","Sheet4")).to be_a(Float) }
      end

      context 'float with point' do
        it { expect(subject.cell(3,"C","Sheet4")).to eq(1234.12) }
        it { expect(subject.cell(3,"C","Sheet4")).to be_a(Float) }
      end
    end

    context 'file path is a Pathname' do
      subject do
        Roo::OpenOffice.new(Pathname.new('test/files/numbers1.ods'))
      end

      it 'creates an instance' do
        expect(subject).to be_a(Roo::OpenOffice)
      end
    end

  end

  # OpenOffice is an alias of LibreOffice. See libreoffice_spec.
end
