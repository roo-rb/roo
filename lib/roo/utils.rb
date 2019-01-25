# frozen_string_literal: true

module Roo
  module Utils
    extend self

    LETTERS = ('A'..'Z').to_a

    def extract_coordinate(s)
      num = letter_num = 0
      num_only = false

      s.each_byte do |b|
        if !num_only && (index = char_index(b))
          letter_num *= 26
          letter_num += index
        elsif index = num_index(b)
          num_only = true
          num *= 10
          num += index
        else
          fail ArgumentError
        end
      end
      fail ArgumentError if letter_num == 0 || !num_only

      Excelx::Coordinate.new(num, letter_num)
    end

    alias_method :ref_to_key, :extract_coordinate

    def split_coordinate(str)
      warn "[DEPRECATION] `Roo::Utils.split_coordinate` is deprecated.  Please use `Roo::Utils.extract_coordinate` instead."
      extract_coordinate(str)
    end



    def split_coord(str)
      coord = extract_coordinate(str)
      [number_to_letter(coord.column), coord.row]
    end

    # convert a number to something like 'AB' (1 => 'A', 2 => 'B', ...)
    def number_to_letter(num)
      result = +""

      until num.zero?
        num, index = (num - 1).divmod(26)
        result.prepend(LETTERS[index])
      end

      result
    end

    def letter_to_number(letters)
      @letter_to_number ||= {}
      @letter_to_number[letters] ||= begin
         result = 0

         # :bytes method returns an enumerator in 1.9.3 and an array in 2.0+
         letters.bytes.to_a.map{|b| b > 96 ? b - 96 : b - 64 }.reverse.each_with_index{ |num, i| result += num * 26 ** i }

         result
      end
    end

    # Compute upper bound for cells in a given cell range.
    def num_cells_in_range(str)
      cells = str.split(':')
      return 1 if cells.count == 1
      raise ArgumentError.new("invalid range string: #{str}. Supported range format 'A1:B2'") if cells.count != 2
      x1, y1 = extract_coordinate(cells[0])
      x2, y2 = extract_coordinate(cells[1])
      (x2 - (x1 - 1)) * (y2 - (y1 - 1))
    end

    def coordinates_in_range(str)
      return to_enum(:coordinates_in_range, str) unless block_given?
      coordinates = str.split(":", 2).map! { |s| extract_coordinate s }

      case coordinates.size
      when 1
        yield coordinates[0]
      when 2
        tl, br = coordinates
        rows = tl.row..br.row
        cols = tl.column..br.column
        rows.each do |row|
          cols.each do |column|
            yield Excelx::Coordinate.new(row, column)
          end
        end
      end
    end

    def load_xml(path)
      ::File.open(path, 'rb') do |file|
        ::Nokogiri::XML(file)
      end
    end

    # Yield each element of a given type ('row', 'c', etc.) to caller
    def each_element(path, elements)
      elements = Array(elements)
      Nokogiri::XML::Reader(::File.open(path, 'rb'), nil, nil, Nokogiri::XML::ParseOptions::NOBLANKS).each do |node|
        next unless node.node_type == Nokogiri::XML::Reader::TYPE_ELEMENT && elements.include?(node.name)
        yield Nokogiri::XML(node.outer_xml).root if block_given?
      end
    end

    private

    def char_index(byte)
      if byte >= 65 && byte <= 90
        byte - 64
      elsif byte >= 97 && byte <= 122
        byte - 96
      end
    end

    def num_index(byte)
      if byte >= 48 && byte <= 57
        byte - 48
      end
    end
  end
end
