require 'spec_helper'

describe Roo do

  describe "version" do
    it{ expect(Roo::VERSION).to match String }
  end
end
