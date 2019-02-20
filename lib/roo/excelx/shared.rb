module Roo
  class Excelx
    # Public:  Shared class for allowing sheets to share data. This should
    #          reduce memory usage and reduce the number of objects being passed
    #          to various inititializers.
    class Shared
      attr_accessor :comments_files, :sheet_files, :rels_files, :image_rels, :image_files
      def initialize(dir, options = {})
        @dir = dir
        @comments_files = []
        @sheet_files = []
        @rels_files = []
        @options = options
        @image_rels = []
        @image_files = []
      end

      def styles
        @styles ||= Styles.new(File.join(@dir, 'roo_styles.xml'))
      end

      def shared_strings
        @shared_strings ||= SharedStrings.new(File.join(@dir, 'roo_sharedStrings.xml'), @options)
      end

      def workbook
        @workbook ||= Workbook.new(File.join(@dir, 'roo_workbook.xml'))
      end

      def base_date
        workbook.base_date
      end

      def base_timestamp
        workbook.base_timestamp
      end
    end
  end
end
