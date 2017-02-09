module Roo
  module Formatters
    module YAML
      # returns a rectangular area (default: all cells) as yaml-output
      # you can add additional attributes with the prefix parameter like:
      # oo.to_yaml({"file"=>"flightdata_2007-06-26", "sheet" => "1"})
      def to_yaml(prefix = {}, from_row = nil, from_column = nil, to_row = nil, to_column = nil, sheet = default_sheet)
        # return an empty string if there is no first_row, i.e. the sheet is empty
        return "" unless first_row

        from_row ||= first_row(sheet)
        to_row ||= last_row(sheet)
        from_column ||= first_column(sheet)
        to_column ||= last_column(sheet)

        result = "--- \n"
        from_row.upto(to_row) do |row|
          from_column.upto(to_column) do |col|
            next if empty?(row, col, sheet)

            result << "cell_#{row}_#{col}: \n"
            prefix.each do|k, v|
              result << "  #{k}: #{v} \n"
            end
            result << "  row: #{row} \n"
            result << "  col: #{col} \n"
            result << "  celltype: #{celltype(row, col, sheet)} \n"
            value = cell(row, col, sheet)
            if celltype(row, col, sheet) == :time
              value = integer_to_timestring(value)
            end
            result << "  value: #{value} \n"
          end
        end

        result
      end
    end
  end
end
