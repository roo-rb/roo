require 'roo/excelx/extractor'

module Roo
  class Excelx
    class Workbook < Excelx::Extractor
      class Label
        attr_reader :sheet, :row, :col, :name

        def initialize(name, sheet, row, col)
          @name = name
          @sheet = sheet
          @row = row.to_i
          @col = ::Roo::Utils.letter_to_number(col)
        end

        def key
          [@row, @col]
        end
      end

      def initialize(path)
        super
        fail ArgumentError, 'missing required workbook file' unless doc_exists?
      end

      def sheets
        doc.xpath('//sheet')
      end

      # aka labels
      def defined_names
        Hash[doc.xpath('//definedName').map do |defined_name|
          # "Sheet1!$C$5"
          sheet, coordinates = defined_name.text.split('!$', 2)
          col, row = coordinates.split('$')
          name = defined_name['name']
          [name, Label.new(name, sheet, row, col)]
        end]
      end

      EPOCH_1900 = Date.new(1900, 1, 1).freeze
      EPOCH_1904 = Date.new(1904, 1, 1).freeze

      def base_date
        @base_date ||=
          begin
            # Default to 1900 (minus one day due to excel quirk) but use 1904 if
            # it's set in the Workbook's workbookPr
            # http://msdn.microsoft.com/en-us/library/ff530155(v=office.12).aspx
            doc.css('workbookPr[date1904]').any? {|workbookPr| workbookPr['date1904'] =~ /true|1/i} ?
              EPOCH_1904 :
              EPOCH_1900
          end
      end
    end
  end
end
