require 'spec_helper'

describe Roo::Excelx do
  describe '.new' do
    subject {
      Roo::Excelx.new('test/files/numbers1.xlsx')
    }

    it 'creates an instance' do
      expect(subject).to be_a(Roo::Excelx)
    end
  end

  describe '.new with streaming' do
    subject {
      Roo::Excelx.new('test/files/numbers1withnull.xlsx', :minimal_load => true)
    }

    it 'has pad_cell set to true by default' do
        # last row should have two nil values
        # on in the first col and one in the fourth col

        rows = []

        subject.each_row_streaming do |row|
            rows << row.map do |c| 
                # can't call .value on a padded value (it's nil)
                c.value rescue nil
            end
        end

        # were padding empty cells (should be 2 total)
        expect(rows.first[1]).to be_nil
        expect(rows.first[3]).to be_nil
        expect(rows.first[2]).not_to be_nil
        
    end

    it 'has pad_cell set to true by default' do
        # last row should have two nil values
        # on in the first col and one in the fourth col

        rows = []

        subject.each_row_streaming(:pad_cells=>false) do |row|
            rows << row.map do |c| 
                # can't call .value on a padded value (it's nil)
                c.value rescue nil
            end
        end

        # we are not padding empty cells
        expect(rows.first[1]).not_to be_nil
        
    end


  end

end
