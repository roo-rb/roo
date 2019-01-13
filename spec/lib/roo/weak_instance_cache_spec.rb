require 'spec_helper'

if RUBY_PLATFORM == "java"
  require 'java'
  java_import 'java.lang.System'
end

describe Roo::Helpers::WeakInstanceCache do
  let(:klass) do
    Class.new do
      include Roo::Helpers::WeakInstanceCache

      def memoized_data
        instance_cache(:@memoized_data) do
          "Some Costly Operation #{rand(1000)}" * 1_000
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
    expect(subject.instance_variables).to_not include(:@memoized_data)
    GC.disable
    subject.memoized_data && nil
    expect(subject.instance_variables).to include(:@memoized_data)

    force_gc
    expect(subject.instance_variables).to_not include(:@memoized_data)
    GC.disable
    subject.memoized_data && nil
    expect(subject.instance_variables).to include(:@memoized_data)
  end

  it 'must remove instance variable' do
    expect(subject.instance_variables).to_not include(:@memoized_data)
    GC.disable
    subject.memoized_data && nil
    expect(subject.instance_variables).to include(:@memoized_data)

    force_gc
    expect(subject.instance_variables).to_not include(:@memoized_data)
  end

  context '#inspect must not raise' do
    it 'before calculation' do
      expect{subject.inspect}.to_not raise_error
    end
    it 'after calculation' do
      GC.disable
      subject.memoized_data && nil
      expect{subject.inspect}.to_not raise_error
      expect(subject.inspect).to include("Some Costly Operation")
      force_gc
    end
    it 'after GC' do
      subject.memoized_data && nil
      force_gc
      expect(subject.instance_variables).to_not include(:@memoized_data)
      expect{subject.inspect}.to_not raise_error
      expect(subject.inspect).to_not include("Some Costly Operation")
    end
  end

  if RUBY_PLATFORM == "java"
    def force_gc
      System.gc
      sleep(0.1)
    end
  else
    def force_gc
      GC.start(full_mark: true, immediate_sweep: true)
      sleep(0.1)
      GC.start(full_mark: true, immediate_sweep: true)
    end
  end
end