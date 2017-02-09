require 'uri'

module Roo
  class Link < String
    # FIXME: Roo::Link inherits from String. A link cell is_a?(Roo::Link). **It is
    #        the only situation where a cells `value` is always a String**. Link
    #        cells have a nifty `to_uri` method, but this method isn't easily
    #        reached. (e.g. `sheet.sheet_for(nil).cells[[row,column]]).value.to_uri`;
    #        `sheet.hyperlink(row, column)` doesn't use `to_uri`).
    #
    #        1. Add different types of links (String, Numeric, Date, DateTime, etc.)
    #        2. Remove Roo::Link.
    #        3. Don't inherit the string and pass the cell's value.
    #
    #        I don't know the historical reasons for the Roo::Link, but right now
    #        it seems uneccessary. I'm in favor of keeping it just in case.
    #
    #        I'm also in favor of passing the cell's value to Roo::Link. The
    #        cell.value's class would still be Roo::Link, but the value itself
    #        would depend on what type of cell it is (Numeric, Date, etc.).
    #
    attr_reader :href
    alias_method :url, :href

    def initialize(href = '', text = href)
      super(text)
      @href = href
    end

    def to_uri
      URI.parse href
    end
  end
end
