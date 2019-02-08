require 'roo/excelx/extractor'
require 'roo/excelx/drawing/base'
require 'roo/excelx/drawing/checkbox'
require 'roo/excelx/drawing/unknown'

module Roo
  class Excelx
    class Drawings < Excelx::Extractor
      # Returns: Hash { id1: extracted_file_name1 },
      # Example: { "rId1"=>"roo_media_image1.png",
      #            "rId2"=>"roo_media_image2.png",
      #            "rId3"=>"roo_media_image3.png" }
      def list
        @drawings ||= extract_drawings
      end

      private

      def extract_drawings
        return {} unless doc_exists?

        doc.xpath('//shape').each_with_object({}) do |shape, hash|
          client_data = shape.xpath('./ClientData').first
          anchors = client_data.xpath('./Anchor').text
          coords = Utils.anchor_to_coordinates(anchors)

          hash[coords] = create_drawing(
            client_data['ObjectType'],
            client_data,
            coords
          )
        end
      end


      def create_drawing(type, client_data, coords)
        case type
        when 'Checkbox'
          Drawing::Checkbox.new(type, client_data.xpath('./Checked').any?, coords)
        else
          Drawing::Unknown.new(type, nil, coords)
        end
      end
    end
  end
end
