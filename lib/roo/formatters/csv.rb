module Roo
  module Formatters
    module CSV
      def to_csv(filename = nil, separator = ",", sheet = default_sheet)
        if filename
          File.open(filename, "w") do |file|
            write_csv_content(file, sheet, separator)
          end
          true
        else
          sio = ::StringIO.new
          write_csv_content(sio, sheet, separator)
          sio.rewind
          sio.read
        end
      end

      private

      # Write all cells to the csv file. File can be a filename or nil. If the
      # file argument is nil the output goes to STDOUT
      def write_csv_content(file = nil, sheet = nil, separator = ",")
        file ||= STDOUT
        return unless first_row(sheet) # The sheet is empty

        1.upto(last_row(sheet)) do |row|
          1.upto(last_column(sheet)) do |col|
            # TODO: use CSV.generate_line
            file.print(separator) if col > 1
            file.print cell_to_csv(row, col, sheet)
          end
          file.print("\n")
        end
      end

      # The content of a cell in the csv output
      def cell_to_csv(row, col, sheet)
        return "" if empty?(row, col, sheet)

        onecell = cell(row, col, sheet)

        case celltype(row, col, sheet)
        when :string
          %("#{onecell.gsub('"', '""')}") unless onecell.empty?
        when :boolean
          # TODO: this only works for excelx
          onecell = self.sheet_for(sheet).cells[[row, col]].formatted_value
          %("#{onecell.gsub('"', '""').downcase}")
        when :float, :percentage
          if onecell == onecell.to_i
            onecell.to_i.to_s
          else
            onecell.to_s
          end
        when :formula
          case onecell
          when String
            %("#{onecell.gsub('"', '""')}") unless onecell.empty?
          when Integer
            onecell.to_s
          when Float
            if onecell == onecell.to_i
              onecell.to_i.to_s
            else
              onecell.to_s
            end
          when Date, DateTime, TrueClass, FalseClass
            onecell.to_s
          else
            fail "unhandled onecell-class #{onecell.class}"
          end
        when :date, :datetime
          onecell.to_s
        when :time
          integer_to_timestring(onecell)
        when :link
          %("#{onecell.url.gsub('"', '""')}")
        else
          fail "unhandled celltype #{celltype(row, col, sheet)}"
        end || ""
      end
    end
  end
end
