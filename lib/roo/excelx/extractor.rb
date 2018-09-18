# frozen_string_literal: true

require "roo/helpers/weak_instance_cache"

module Roo
  class Excelx
    class Extractor
      include Roo::Helpers::WeakInstanceCache

      COMMON_STRINGS = {
        t: "t",
        r: "r",
        s: "s",
        ref: "ref",
        html_tag_open: "<html>",
        html_tag_closed: "</html>"
      }

      def initialize(path, options = {})
        @path = path
        @options = options
      end

      private

      def doc
        instance_cache(:@doc) do
          raise FileNotFound, "#{@path} file not found" unless doc_exists?

          ::Roo::Utils.load_xml(@path).remove_namespaces!
        end
      end

      def doc_exists?
        @path && File.exist?(@path)
      end
    end
  end
end
