# frozen_string_literal: true

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

      def include_type?(type)
        to_a.any? do |_, rel|
          rel["Type"]&.include? type
        end
      end

      def target(type)
        to_a.map do |_, rel|
          rel["Target"] if rel["Type"]&.include? type
        end.compact.first
      end

      private

      def extract_relationships
        return {} unless doc_exists?

        doc.xpath('/Relationships/Relationship').each_with_object({}) do |rel, hash|
          hash[rel['Id']] = rel
        end
      end
    end
  end
end
