require 'spec_helper'

describe Roo::Excelx::Format do
  describe '#to_type' do
    FORMATS = {
      'General' => :float,
      '0' => :float,
      '0.00' => :float,
      '#,##0' => :float,
      '#,##0.00' => :float,
      '0%' => :percentage,
      '0.00%' => :percentage,
      '0.00E+00' => :float,
      '# ?/?' => :float, # ??? TODO:
      '# ??/??' => :float, # ??? TODO:
      'mm-dd-yy' => :date,
      'd-mmm-yy' => :date,
      'd-mmm' => :date,
      'mmm-yy' => :date,
      'h:mm AM/PM' => :date,
      'h:mm:ss AM/PM' => :date,
      'h:mm' => :time,
      'h:mm:ss' => :time,
      'm/d/yy h:mm' => :datetime,
      '#,##0 ;(#,##0)' => :float,
      '#,##0 ;[Red](#,##0)' => :float,
      '#,##0.00;(#,##0.00)' => :float,
      '#,##0.00;[Red](#,##0.00)' => :float,
      '#,##0_);[Red](#,##0)' => :float,
      'mm:ss' => :time,
      '[h]:mm:ss' => :time,
      'mmss.0' => :time,
      '##0.0E+0' => :float,
      '@' => :float,
      #-- zusaetzliche Formate, die nicht standardmaessig definiert sind:
      'yyyy\\-mm\\-dd' => :date,
      'dd/mm/yy' => :date,
      'hh:mm:ss' => :time,
      'dd/mm/yy\\ hh:mm' => :datetime,
      'dd/mmm/yy\\ hh:mm' => :datetime,
      'dd/mmm/yy' => :date, # 2011-05-21
      'yyyy-mm-dd' => :date, # 2011-09-16
      'yyyy-mm-dd;@' => :date,
      '#0_);[Red]\(0\)' => :float
    }.each do |format, type|
      it "translates #{format} to #{type}" do
        expect(Roo::Excelx::Format.to_type(format)).to eq(type)
      end
    end
  end
end
