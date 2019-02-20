require 'roo/font'
require 'roo/excelx/extractor'

module Roo
  class Excelx
    class Styles < Excelx::Extractor
      # convert internal excelx attribute to a format
      def style_format(style)
        id = num_fmt_ids[style.to_i]
        num_fmts[id] || Excelx::Format::STANDARD_FORMATS[id.to_i]
      end

      def definitions
        @definitions ||= extract_definitions
      end

      private

      def num_fmt_ids
        @num_fmt_ids ||= extract_num_fmt_ids
      end

      def num_fmts
        @num_fmts ||= extract_num_fmts
      end

      def fonts
        @fonts ||= extract_fonts
      end

      def extract_definitions
        doc.xpath('//cellXfs').flat_map do |xfs|
          xfs.children.map do |xf|
            fonts[xf['fontId'].to_i]
          end
        end
      end

      def extract_fonts
        doc.xpath('//fonts/font').map do |font_el|
          Font.new.tap do |font|
            font.bold = !font_el.xpath('./b').empty?
            font.italic = !font_el.xpath('./i').empty?
            font.underline = !font_el.xpath('./u').empty?
          end
        end
      end

      def extract_num_fmt_ids
        doc.xpath('//cellXfs').flat_map do |xfs|
          xfs.children.map do |xf|
            xf['numFmtId']
          end
        end.compact
      end

      def extract_num_fmts
        doc.xpath('//numFmt').each_with_object({}) do |num_fmt, hash|
          hash[num_fmt['numFmtId']] = num_fmt['formatCode']
        end
      end
    end
  end
end
