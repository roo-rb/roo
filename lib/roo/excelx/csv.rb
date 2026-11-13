# frozen_string_literal: true

require "English"

module Roo
  class Excelx
    module SparseCSVWriter
      private

      def write_csv_content(file = nil, sheet = nil, separator = ",")
        file ||= STDOUT
        return super unless csv_walk_compatible?(separator)
        return super unless [File, StringIO].include?(file.class) && file.singleton_methods.empty?
        return unless first_row(sheet)

        target = sheet_for(sheet)
        return super unless target.instance_of?(Sheet) && target.singleton_methods.empty?
        cells = target.cells

        rows = last_row(sheet)
        width = last_column(sheet)
        return super unless rows.is_a?(Integer) && width.is_a?(Integer) && width.positive?
        return super if cells.size * 2 >= rows * width
        return super unless csv_cells_compatible?(cells)

        separator = separator.to_s
        runs = Hash.new { |cache, count| cache[count] = separator * count }
        blank = "#{runs[width - 1]}\n"
        line = +""
        line_row = column_written = 0
        each_csv_coordinate(cells, rows, width) do |row, column|
          if row != line_row
            file.print(line << runs[width - 1 - column_written] << "\n") if line_row.positive?
            (row - line_row - 1).times { file.print(blank) }
            line.clear
            line_row = row
            column_written = 0
          end
          line << runs[column - 1 - column_written]
          column_written = column - 1
          begin
            field = cell_to_csv(row, column, sheet)
          rescue StandardError
            file.print(line)
            raise
          end
          if field.ascii_only? || field.encoding == Encoding::UTF_8
            line << field
          else
            file.print(line)
            line.clear
            file.print(field)
          end
        end
        file.print(line << runs[width - 1 - column_written] << "\n") if line_row.positive?
        (line_row + 1).upto(rows) { file.print(blank) }
      end

      def each_csv_coordinate(cells, rows, columns)
        previous_row = previous_column = 0
        ordered = cells.each_key.all? do |key|
          next false unless key.is_a?(Array) && key.size == 2 && key.all?(Integer)

          row, column = key
          next false if row < previous_row || (row == previous_row && column < previous_column)

          previous_row, previous_column = row, column
          true
        end
        if ordered
          cells.each_key do |row, column|
            break if row > rows

            yield row, column if row.positive? && column.between?(1, columns)
          end
        else
          keys = cells.keys.select do |key|
            key.is_a?(Array) && key.size == 2 && key.all?(Integer) &&
              key[0].between?(1, rows) && key[1].between?(1, columns)
          end.sort!
          keys.each { |row, column| yield row, column }
        end
      end

      def csv_walk_compatible?(separator)
        instance_of?(Roo::Excelx) && $OUTPUT_RECORD_SEPARATOR.nil? &&
          method(:normalize).owner == Roo::Base &&
          [String, Symbol, NilClass].include?(separator.class) &&
          (separator.to_s.ascii_only? || separator.to_s.encoding == Encoding::UTF_8) &&
          %i[cell celltype empty? sheet_for first_row last_row last_column].all? do |name|
            method(name).owner == Roo::Excelx
          end && method(:cell_to_csv).owner == Roo::Formatters::CSV
      end

      def csv_cells_compatible?(cells)
        return false unless cells.instance_of?(Hash) && cells.default.nil? && cells.default_proc.nil?

        classes = [Cell::String, Cell::Boolean, Cell::Number, Cell::Date, Cell::DateTime, Cell::Time, Cell::Empty]
        values = [String, Roo::Link, Integer, Float, ::Date, ::DateTime, TrueClass, FalseClass, NilClass]
        cells.each_value.all? do |cell|
          next true if cell.nil?
          next false unless classes.include?(cell.class) && cell.singleton_methods.empty?

          value = cell.instance_variable_get(:@value)
          next false unless values.include?(value.class)
          next false if (value.is_a?(::Date) || value.is_a?(String)) && !value.singleton_methods.empty?
          if value.instance_of?(Roo::Link)
            href = value.instance_variable_get(:@href)
            next false unless href.instance_of?(String) && href.singleton_methods.empty?
          end

          true
        end
      end
    end
  end
end
