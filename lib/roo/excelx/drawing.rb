require 'roo/excelx/extractor'

module Roo
  class Excelx
    class Drawing < Excelx::Extractor

      # Returns: Hash { id1: cell_coordinates },
      # Example: { "rId1"=> { from_col: 2, from_row: 3, to_col: 2, to_row: 3 },
      #            "rId2"=> { from_col: 2, from_row: 4, to_col: 2, to_row: 4 },
      #            "rId3"=> { from_col: 2, from_row: 5, to_col: 2, to_row: 5 } }
      #
      def list
        @image_coordinates ||= extract_image_coordinates
      end

      private

      def extract_image_coordinates
        return {} unless doc_exists?
        data = Hash.new { |hash, key| hash[key] = [] }

        # Loop through all twoCellAnchor elements and extract the information
        doc.xpath('//twoCellAnchor').each do |anchor|
          # Extract the row and column numbers
          from_col = anchor.at_xpath('./from/col').text.to_i
          from_row = anchor.at_xpath('./from/row').text.to_i
          to_col = anchor.at_xpath('./to/col').text.to_i
          to_row = anchor.at_xpath('./to/row').text.to_i

          # Extract the rId attribute from the blip element if present, if not ignore anchor element
          if anchor.at_xpath('./pic/blipFill/blip')
            r_id = anchor.at_xpath('./pic/blipFill/blip')['embed']

            # Store the extracted information in the data hash
            data[r_id] << { from_col: from_col, from_row: from_row, to_col: to_col, to_row: to_row }
          end
        end

        # Loop through all oneCellAnchor elements and extract the information
        doc.xpath('//oneCellAnchor').each do |anchor|
          # Extract the row and column numbers
          from_col = anchor.at_xpath('./from/col')&.text&.to_i
          from_row = anchor.at_xpath('./from/row')&.text&.to_i

          # Extract the rId attribute from the blip element if present, if not ignore anchor element
          if anchor.at_xpath('./pic/blipFill/blip')
            r_id = anchor.at_xpath('./pic/blipFill/blip')['embed']

            # Store the extracted information in the data hash
            data[r_id] << { from_col: from_col, from_row: from_row }
          end
        end

        data
      end
    end
  end
end
