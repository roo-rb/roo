require 'roo/excelx/extractor'

module Roo
  class Excelx
    class Relationships < Excelx::Extractor
      def [](index)
        to_a[index]
      end

      def to_a
        @relationships ||= extract_relationships
      end

      private

      def extract_relationships
        return [] unless doc_exists?

        doc.xpath('/Relationships/Relationship').each_with_object({}) do |rel, hash|
          hash[rel.attribute('Id').text] = rel
        end
      end
    end
  end
end
