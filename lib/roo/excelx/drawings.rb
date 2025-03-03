require 'roo/excelx/extractor'
require 'roo/excelx/drawing/base'
require 'roo/excelx/drawing/checkbox'
require 'roo/excelx/drawing/unknown'

module Roo
  class Excelx
    class Drawings < Excelx::Extractor
      # Returns: Hash { [col, row]: drawing },
      # Example: { [10, 3]=>#<Roo::Excelx::Drawing::Checkbox>,
      #            [19, 7]=>#<Roo::Excelx::Drawing::Checkbox>,
      #            [1, 14]=>#<Roo::Excelx::Drawing::Unknown> }
      def list
        @drawings ||= extract_drawings
      end

      private

      def extract_drawings
        return {} unless doc_exists?

        doc.xpath('//shape/ClientData').each_with_object({}) do |client_data, hash|
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
