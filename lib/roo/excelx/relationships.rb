require 'roo/excelx/extractor'

module Roo
  class Excelx
    class Relationships < Excelx::Extractor
      def [](index)
        to_h[index]
      end

      def to_h
        @relationships ||= extract_relationships
      end

      private

      def extract_relationships
        return {} unless doc_exists?

        Hash[doc.xpath('/Relationships/Relationship').map do |rel|
          [rel.attribute('Id').text, rel]
        end]
      end
    end
  end
end
