require 'roo/excelx/extractor'

module Roo
  class Excelx
    class Drawing < Excelx::Extractor

      # Returns: Hash { id1: cell_coordinates },
      # Example: { "rId1"=> [2,4],
      #            "rId2"=> [4,3],
      #            "rId3"=> [5,4] }
      def list
        @image_coordinates ||= extract_image_coordinates
      end

      private

      def extract_image_coordinates
        return {} unless doc_exists?
        data = {}

        # Loop through all twoCellAnchor elements and extract the information
        doc.xpath('//twoCellAnchor').each do |anchor|
          # Extract the row and column numbers
          from_col = anchor.at_xpath('./from/col').text.to_i
          from_row = anchor.at_xpath('./from/row').text.to_i
          to_col = anchor.at_xpath('./to/col').text.to_i
          to_row = anchor.at_xpath('./to/row').text.to_i

          # Extract the rId attribute from the blip element
          if anchor.at_xpath('./pic/blipFill/blip')
            r_id = anchor.at_xpath('./pic/blipFill/blip')['embed']

            # Store the extracted information in the data hash
            data[r_id] = { from_col: from_col, from_row: from_row, to_col: to_col, to_row: to_row }
          end
        end

        data
      end
    end
  end
end
