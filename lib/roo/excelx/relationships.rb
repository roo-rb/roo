require 'roo/excelx/extractor'

module Roo
  class Excelx::Relationships < Excelx::Extractor
    def [](index)
      to_a[index]
    end

    def to_a
      @relationships ||= extract_relationships
    end

    private

    def extract_relationships
      if doc_exists?
        Hash[doc.xpath("/Relationships/Relationship").map do |rel|
          [rel.attribute('Id').text, rel]
        end]
      else
        []
      end
    end

  end
end
