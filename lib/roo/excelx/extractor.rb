module Roo
  class Excelx
    class Extractor
      def initialize(path, options)
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
