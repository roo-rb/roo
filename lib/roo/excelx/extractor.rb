module Roo
  class Excelx
    class Extractor
      def initialize(path)
        @path = path
      end

      private

      def doc
        @doc ||=
        if doc_exists?
          ::Roo::Utils.load_xml(@path).remove_namespaces!
        end
      end

      def doc_exists?
        @path && File.exist?(@path)
      end
    end
  end
end
