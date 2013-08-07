require 'spec_helper'

describe Roo::Google do
  let(:key) { nil }

  describe '.new' do
    subject {
      Roo::Google.new(key)
    }

    it 'creates an instance' do
      pending "we need to come up with a way to test google that can be used by any developer"
      expect(subject).to be_a(Roo::Google)
    end
  end
end
