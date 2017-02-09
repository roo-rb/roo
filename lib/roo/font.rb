module Roo
  class Font
    attr_accessor :bold, :italic, :underline

    def bold?
      @bold
    end

    def italic?
      @italic
    end

    def underline?
      @underline
    end
  end
end
