require 'spec_helper'

RSpec.describe Roo::Excelx::Cell::Number do

  describe '#initialize' do
    let(:formula)     { nil }
    let(:excelx_type) { [:numeric_or_formula, format] }
    let(:format)      { "#,##0.00" }
    let(:style)       { 1 }
    let(:link)        { nil }
    let(:coordinate)  { [1,1] }

    let(:number)  { Roo::Excelx::Cell::Number.new(value, formula, excelx_type, style, link, coordinate)}

    context 'with an actual value' do
      let(:value)       { 0.1 }
      it 'creates the object and parses the value' do
        expect(number.value).to eq 0.1
      end
    end

    context 'with an empty value' do
      let(:value)       { "" }
      it 'creates the object with a nil value' do
        expect(number.value).to eq nil
      end
    end
  end

  describe '#formatted_value' do
    let(:formula)     { nil }
    let(:excelx_type) { [:numeric_or_formula, format] }
    let(:format)      { "#,##0.00" }
    let(:style)       { 1 }
    let(:link)        { nil }
    let(:coordinate)  { [1,1] }

    let(:number)  { Roo::Excelx::Cell::Number.new(value, formula, excelx_type, style, link, coordinate)}

    context 'with an actual value' do
      let(:value)       { 0.1 }
      it 'creates the object and parses the value' do
        expect(number.formatted_value).to eq '0.10'
      end
    end

    context 'with an empty value' do
      let(:value)       { "" }
      it 'creates the object with a nil value' do
        expect(number.formatted_value).to eq ''
      end
    end

  end
end
