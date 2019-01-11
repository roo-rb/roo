require 'spec_helper'

describe Roo::Helpers::WeakInstanceCache do
  let(:klass) do
    Class.new do
      include Roo::Helpers::WeakInstanceCache

      def memoized_data
        instance_cache(:@memoized_data) do
          "Some Costly Operation" * 1_000
        end
      end
    end
  end

  subject do
    klass.new
  end

  it 'should be lazy' do
    expect(subject.instance_variables).to_not include(:@memoized_data)
    data = subject.memoized_data
    expect(subject.instance_variables).to include(:@memoized_data)
  end


  it 'should be memoized' do
    data = subject.memoized_data
    expect(subject.memoized_data).to equal(data)
  end

  it 'should recalculate after GC' do
    GC.disable
    original_id = subject.memoized_data.object_id
    expect(subject.memoized_data.object_id).to eq(original_id)

    GC.start
    expect(subject.memoized_data.object_id).not_to eq(original_id)
  end

  it 'must remove instance variable' do
    expect(subject.instance_variables).to_not include(:@memoized_data)
    GC.disable
    subject.memoized_data
    expect(subject.instance_variables).to include(:@memoized_data)

    GC.start
    expect(subject.instance_variables).to_not include(:@memoized_data)
  end

  context '#inspect must not raise' do
    it 'before calculation' do
      expect{subject.inspect}.to_not raise_error
    end
    it 'after calculation' do
      GC.disable
      subject.memoized_data
      expect{subject.inspect}.to_not raise_error
      expect(subject.inspect).to include("Some Costly Operation")
      GC.start
    end
    it 'after GC' do
      subject.memoized_data
      GC.start
      expect{subject.inspect}.to_not raise_error
      expect(subject.inspect).to_not include("Some Costly Operation")
    end
  end
end