module Roo
  module Utils
    extend self

    LETTERS = ('A'..'Z').to_a

    def split_coordinate(str)
      @split_coordinate ||= {}

      @split_coordinate[str] ||= begin
        letter, number = split_coord(str)
        x = letter_to_number(letter)
        y = number
        [y, x]
      end
    end

    alias_method :ref_to_key, :split_coordinate

    def split_coord(s)
      if s =~ /([a-zA-Z]+)([0-9]+)/
        letter = Regexp.last_match[1]
        number = Regexp.last_match[2].to_i
      else
        fail ArgumentError
      end
      [letter, number]
    end

    # convert a number to something like 'AB' (1 => 'A', 2 => 'B', ...)
    def number_to_letter(num)
      result = ""

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
      x1, y1 = split_coordinate(cells[0])
      x2, y2 = split_coordinate(cells[1])
      (x2 - (x1 - 1)) * (y2 - (y1 - 1))
    end

    def load_xml(path)
      ::File.open(path, 'rb') do |file|
        ::Nokogiri::XML(file)
      end
    end

    # Yield each element of a given type ('row', 'c', etc.) to caller
    def each_element(path, elements)
      Nokogiri::XML::Reader(::File.open(path, 'rb'), nil, nil, Nokogiri::XML::ParseOptions::NOBLANKS).each do |node|
        next unless node.node_type == Nokogiri::XML::Reader::TYPE_ELEMENT && Array(elements).include?(node.name)
        yield Nokogiri::XML(node.outer_xml).root if block_given?
      end
    end
  end
end
