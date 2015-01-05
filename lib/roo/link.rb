module Roo
  class Link < String
    attr_reader :href
    alias :url :href

    def initialize(href='', text=href)
      super(text)
      @href = href
    end

    def to_uri
      URI.parse href
    end
  end
end
