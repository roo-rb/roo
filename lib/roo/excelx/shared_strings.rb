require 'roo/excelx/extractor'

module Roo
  class Excelx
    class SharedStrings < Excelx::Extractor
      def [](index)
        to_a[index]
      end

      def to_a
        @array ||= extract_shared_strings
      end

      private

      def extract_shared_strings
        return [] unless doc_exists?

        # read the shared strings xml document
        doc.xpath('/sst/si').map do |si|
          shared_string = ''
          si.children.each do |elem|
            case elem.name
            when 'r'
              elem.children.each do |r_elem|
                shared_string << r_elem.content if r_elem.name == 't'
              end
            when 't'
              shared_string = elem.content
            end
          end
          shared_string
        end
      end
    end
  end
end
