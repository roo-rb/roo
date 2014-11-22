module Roo
  class Excelx::Extractor
    def initialize(path)
      @path = path
    end

    private

    def doc
      @doc ||=
        if doc_exists?
          Roo::Excelx.load_xml(@path)
        end
    end

    def doc_exists?
      File.exist?(@path)
    end
  end
end
