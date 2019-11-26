require 'spec_helper'

describe Roo::Excelx do
  subject { Roo::Excelx.new('test/files/strict.xlsx') }

  example '#sheets' do
    expect(subject.sheets).to eq %w(Sheet1 Sheet2)
  end

  example '#sheet' do
    expect(subject.sheet('Sheet1')).to be_a(Roo::Excelx)
  end

  example '#cell' do
    expect(subject.cell(1, 1)).to eq 'Sheet 1'
    expect(subject.cell(1, 1, 'Sheet2')).to eq 'Sheet 2'
  end

  example '#row' do
    expect(subject.row(1)).to eq ['Sheet 1']
    expect(subject.row(1, 'Sheet2')).to eq ['Sheet 2']
  end

  example '#first_row' do
    expect(subject.first_row).to eq 1
    expect(subject.first_row('Sheet2')).to eq 1
  end

  example '#last_row' do
    expect(subject.last_row).to eq 1
    expect(subject.last_row('Sheet2')).to eq 1
  end

  example '#first_column' do
    expect(subject.first_column).to eq 1
    expect(subject.first_column('Sheet2')).to eq 1
  end

  example '#last_column' do
    expect(subject.last_column).to eq 1
    expect(subject.last_column('Sheet2')).to eq 1
  end
end
