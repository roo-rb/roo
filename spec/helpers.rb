module Helpers
  def yaml_entry(row,col,type,value)
    "cell_#{row}_#{col}: \n  row: #{row} \n  col: #{col} \n  celltype: #{type} \n  value: #{value} \n"
  end
end
