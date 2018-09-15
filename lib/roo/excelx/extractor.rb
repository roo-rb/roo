# frozen_string_literal: true

module Roo
  class Excelx
    class Extractor

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
        raise FileNotFound, "#{@path} file not found" unless doc_exists?

        ::Roo::Utils.load_xml(@path).remove_namespaces!
      end

      def doc_exists?
        @path && File.exist?(@path)
      end
    end
  end
end
