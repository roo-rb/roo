module Roo
  module Utils
    extend self
    LETTERS = ('A'..'Z').to_a
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
      if n > 26
        while n % 26 == 0 && n != 0
          letters << 'Z'
          n = ((n - 26) / 26).to_i
        end
        while n > 0
          num     = n % 26
          letters = LETTERS[num - 1] + letters
          n       = (n / 26).to_i
        end
      else
        letters = LETTERS[n - 1]
      end
      letters
    end

    # convert letters like 'AB' to a number ('A' => 1, 'B' => 2, ...)
    def letter_to_number(letters)
      result = 0
      while letters && letters.length > 0
        character = letters[0, 1].upcase
        num       = LETTERS.index(character)
        fail ArgumentError, "invalid column character '#{letters[0, 1]}'" if num.nil?
        num     += 1
        result  = result * 26 + num
        letters = letters[1..-1]
      end
      result
    end

    def load_xml(path)
      ::File.open(path, 'rb') do |file|
        ::Nokogiri::XML(file)
      end
    end
  end
end