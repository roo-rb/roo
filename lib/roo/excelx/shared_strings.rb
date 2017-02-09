require 'roo/excelx/extractor'

module Roo
  class Excelx
    class SharedStrings < Excelx::Extractor

      COMMON_STRINGS = {
        t: "t",
        r: "r",
        html_tag_open: "<html>",
        html_tag_closed: "</html>"
      }

      def [](index)
        to_a[index]
      end

      def to_a
        @array ||= extract_shared_strings
      end

      def to_html
        @html ||= extract_html
      end

      # Use to_html or to_a for html returns
      # See what is happening with commit???
      def use_html?(index)
        to_html[index][/<([biu]|sup|sub)>/]
      end

      private

      def fix_invalid_shared_strings(doc)
        invalid = { '_x000D_'  => "\n" }
        xml = doc.to_s
        return doc unless xml[/#{invalid.keys.join('|')}/]

        ::Nokogiri::XML(xml.gsub(/#{invalid.keys.join('|')}/, invalid))
      end

      def extract_shared_strings
        return [] unless doc_exists?

        document = fix_invalid_shared_strings(doc)
        # read the shared strings xml document
        document.xpath('/sst/si').map do |si|
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

      def extract_html
        return [] unless doc_exists?
        fix_invalid_shared_strings(doc)
        # read the shared strings xml document
        doc.xpath('/sst/si').map do |si|
          html_string = '<html>'
          si.children.each do |elem|
            case elem.name
            when 'r'
              html_string << extract_html_r(elem)
            when 't'
              html_string << elem.content
            end # case elem.name
          end # si.children.each do |elem|
          html_string << '</html>'
        end # doc.xpath('/sst/si').map do |si|
      end # def extract_html

      # The goal of this function is to take the following XML code snippet and create a html tag
      # r_elem ::: XML Element that is in sharedStrings.xml of excel_book.xlsx
      # {code:xml}
      # <r>
      #   <rPr>
      #      <i/>
      #      <b/>
      #      <u/>
      #      <vertAlign val="subscript"/>
      #      <vertAlign val="superscript"/>
      #   </rPr>
      #   <t>TEXT</t>
      # </r>
      # {code}
      #
      # Expected Output ::: "<html><sub|sup><b><i><u>TEXT</u></i></b></sub|/sup></html>"
      def extract_html_r(r_elem)
        str = ''
        xml_elems = {
          sub: false,
          sup: false,
          b:   false,
          i:   false,
          u:   false
        }
        b, i, u, sub, sup = false, false, false, false, false
        r_elem.children.each do |elem|
          case elem.name
          when 'rPr'
            elem.children.each do |rPr_elem|
              rPr_elem_value = rPr_elem.xpath('@val').first.value

              case rPr_elem.name
              when 'b'
                # set formatting for Bold to true if value is one
                xml_elems[:b] = true if rPr_elem_value == '1'
              when 'i'
                # set formatting for Italics to true if value is one
                xml_elems[:i] = true if rPr_elem_value == '1'
              when 'u'
                # set formatting for Underline to true if value is one
                xml_elems[:u] = true if rPr_elem_value == 'single'
              when 'vertAlign'
                # See if the Vertical Alignment is subscript or superscript
                case rPr_elem_value
                when 'subscript'
                  # set formatting for Subscript to true and Superscript to false ... Can't have both
                  xml_elems[:sub] = true
                  xml_elems[:sup] = false
                when 'superscript'
                  # set formatting for Superscript to true and Subscript to false ... Can't have both
                  xml_elems[:sup] = true
                  xml_elems[:sub] = false
                end
              end
            end
          when 't'
            str << create_html(elem.content, xml_elems)
          end
        end
        str
      end # extract_html_r

      # This will return an html string
      def create_html(text, formatting)
        tmp_str = ''
        formatting.each do |elem, val|
          tmp_str << "<#{elem}>" if val
        end
        tmp_str << text
        reverse_format = Hash[formatting.to_a.reverse]
        reverse_format.each do |elem, val|
          tmp_str << "</#{elem}>" if val
        end
        tmp_str
      end
    end # class SharedStrings < Excelx::Extractor
  end # class Excelx
end # module Roo
