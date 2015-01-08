module Roo
  module Utils
    extend self

    LETTERS = ('A'..'Z').to_a
    CODEPOINT_OFFSET = 'A'.ord - 1

    def split_coordinate(str)
      letter, number = split_coord(str)
      x              = letter_to_number(letter)
      y              = number
      [y, x]
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
    def number_to_letter(n)
      letters = ''
      while n > 0
        index = (n - 1) % 26
        letters.prepend(LETTERS[index])
        n = (n - index - 1) / 26
      end
      letters
    end

    # convert letters like 'AB' to a number ('A' => 1, 'B' => 2, ...)
    def letter_to_number(letters)
      letters.upcase.each_codepoint.reduce(0) do |result, codepoint|
        26 * result + codepoint - CODEPOINT_OFFSET
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
