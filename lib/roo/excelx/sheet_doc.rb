# frozen_string_literal: true

require 'forwardable'
require 'roo/excelx/extractor'

module Roo
  class Excelx
    class SheetDoc < Excelx::Extractor
      extend Forwardable
      delegate [:workbook] => :@shared

      def initialize(path, relationships, shared, options = {})
        super(path)
        @shared = shared
        @options = options
        @relationships = relationships
      end

      def cells(relationships)
        @cells ||= extract_cells(relationships)
      end

      def hyperlinks(relationships)
        # If you're sure you're not going to need this hyperlinks you can discard it
        @hyperlinks ||= if @options[:no_hyperlinks] || !relationships.include_type?("hyperlink")
          {}
        else
          extract_hyperlinks(relationships)
        end
      end

      # Get the dimensions for the sheet.
      # This is the upper bound of cells that might
      # be parsed. (the document may be sparse so cell count is only upper bound)
      def dimensions
        @dimensions ||= extract_dimensions
      end

      # Yield each row xml element to caller
      def each_row_streaming(&block)
        Roo::Utils.each_element(@path, 'row', &block)
      end

      # Yield each cell as Excelx::Cell to caller for given
      # row xml
      def each_cell(row_xml)
        return [] unless row_xml
        row_xml.children.each do |cell_element|
          coordinate = ::Roo::Utils.extract_coordinate(cell_element["r"])
          hyperlinks = hyperlinks(@relationships)[coordinate]

          yield cell_from_xml(cell_element, hyperlinks, coordinate)
        end
      end

      private

      def cell_value_type(type, format)
        case type
        when 's'
          :shared
        when 'b'
          :boolean
        when 'str'
          :string
        when 'inlineStr'
          :inlinestr
        else
          Excelx::Format.to_type(format)
        end
      end

      # Internal: Creates a cell based on an XML clell..
      #
      # cell_xml - a Nokogiri::XML::Element. e.g.
      #             <c r="A5" s="2">
      #               <v>22606</v>
      #             </c>
      # hyperlink - a String for the hyperlink for the cell or nil when no
      #             hyperlink is present.
      # coordinate - a Roo::Excelx::Coordinate for the coordinate for the cell
      #              or nil to extract coordinate from cell_xml.
      # empty_cell - an Optional Boolean value.
      #
      # Examples
      #
      #    cells_from_xml(<Nokogiri::XML::Element>, nil, nil)
      #    # => <Excelx::Cell::String>
      #
      # Returns a type of <Excelx::Cell>.
      def cell_from_xml(cell_xml, hyperlink, coordinate, empty_cell=true)
        coordinate ||= ::Roo::Utils.extract_coordinate(cell_xml["r"])
        cell_xml_children = cell_xml.children
        return create_empty_cell(coordinate, empty_cell) if cell_xml_children.empty?

        # NOTE: This is error prone, to_i will silently turn a nil into a 0.
        #       This works by coincidence because Format[0] is General.
        style = cell_xml["s"].to_i
        formula = nil

        cell_xml_children.each do |cell|
          case cell.name
          when 'is'
            content = +""
            cell.children.each do |inline_str|
              if inline_str.name == 't'
                content << inline_str.content
              end
            end
            unless content.empty?
              return Excelx::Cell.cell_class(:string).new(content, formula, style, hyperlink, coordinate)
            end
          when 'f'
            formula = cell.content
          when 'v'
            format = style_format(style)
            value_type = cell_value_type(cell_xml["t"], format)

            return create_cell_from_value(value_type, cell, formula, format, style, hyperlink, coordinate)
          end
        end

        create_empty_cell(coordinate, empty_cell)
      end

      def create_empty_cell(coordinate, empty_cell)
        if empty_cell
          Excelx::Cell::Empty.new(coordinate)
        end
      end

      def create_cell_from_value(value_type, cell, formula, format, style, hyperlink, coordinate)
        # NOTE: format.to_s can replace excelx_type as an argument for
        #       Cell::Time, Cell::DateTime, Cell::Date or Cell::Number, but
        #       it will break some brittle tests.
        excelx_type = [:numeric_or_formula, format.to_s]

        # NOTE: There are only a few situations where value != cell.content
        #       1. when a sharedString is used. value = sharedString;
        #          cell.content = id of sharedString
        #       2. boolean cells: value = 'TRUE' | 'FALSE'; cell.content = '0' | '1';
        #          But a boolean cell should use TRUE|FALSE as the formatted value
        #          and use a Boolean for it's value. Using a Boolean value breaks
        #          Roo::Base#to_csv.
        #       3. formula
        case value_type
        when :shared
          cell_content = cell.content.to_i
          value = shared_strings.use_html?(cell_content) ? shared_strings.to_html[cell_content] : shared_strings[cell_content]
          Excelx::Cell.cell_class(:string).new(value, formula, style, hyperlink, coordinate)
        when :boolean, :string
          value = cell.content
          Excelx::Cell.cell_class(value_type).new(value, formula, style, hyperlink, coordinate)
        when :time, :datetime
          cell_content = cell.content.to_f
          # NOTE: A date will be a whole number. A time will have be > 1. And
          #      in general, a datetime will have decimals. But if the cell is
          #      using a custom format, it's possible to be interpreted incorrectly.
          #      cell_content.to_i == cell_content && standard_style?=> :date
          #
          #      Should check to see if the format is standard or not. If it's a
          #      standard format, than it's a date, otherwise, it is a datetime.
          #      @styles.standard_style?(style_id)
          #      STANDARD_STYLES.keys.include?(style_id.to_i)
          cell_type = if cell_content < 1.0
                        :time
                      elsif (cell_content - cell_content.floor).abs > 0.000001
                        :datetime
                      else
                        :date
                      end
          base_value = cell_type == :date ? base_date : base_timestamp
          Excelx::Cell.cell_class(cell_type).new(cell_content, formula, excelx_type, style, hyperlink, base_value, coordinate)
        when :date
          Excelx::Cell.cell_class(:date).new(cell.content, formula, excelx_type, style, hyperlink, base_date, coordinate)
        else
          Excelx::Cell.cell_class(:number).new(cell.content, formula, excelx_type, style, hyperlink, coordinate)
        end
      end

      def extract_hyperlinks(relationships)
        return {} unless (hyperlinks = doc.xpath('/worksheet/hyperlinks/hyperlink'))

        hyperlinks.each_with_object({}) do |hyperlink, hash|
          if relationship = relationships[hyperlink['id']]
            target_link = relationship['Target']
            target_link += "##{hyperlink['location']}" if hyperlink['location']

            Roo::Utils.coordinates_in_range(hyperlink["ref"].to_s) do |coord|
              hash[coord] = target_link
            end
          end
        end
      end

      def expand_merged_ranges(cells)
        # Extract merged ranges from xml
        merges = {}
        doc.xpath('/worksheet/mergeCells/mergeCell').each do |mergecell_xml|
          tl, br = mergecell_xml["ref"].split(/:/).map { |ref| ::Roo::Utils.ref_to_key(ref) }
          for row in tl[0]..br[0] do
            for col in tl[1]..br[1] do
              next if row == tl[0] && col == tl[1]
              merges[[row, col]] = tl
            end
          end
        end
        # Duplicate value into all cells in merged range
        merges.each do |dst, src|
          cells[dst] = cells[src]
        end
      end

      def extract_cells(relationships)
        extracted_cells = {}
        empty_cell = @options[:empty_cell]

        doc.xpath('/worksheet/sheetData/row/c').each do |cell_xml|
          coordinate = ::Roo::Utils.extract_coordinate(cell_xml["r"])
          cell = cell_from_xml(cell_xml, hyperlinks(relationships)[coordinate], coordinate, empty_cell)
          extracted_cells[coordinate] = cell if cell
        end

        expand_merged_ranges(extracted_cells) if @options[:expand_merged_ranges]

        extracted_cells
      end

      def extract_dimensions
        Roo::Utils.each_element(@path, 'dimension') do |dimension|
          return dimension["ref"]
        end
      end

      def style_format(style)
        @shared.styles.style_format(style)
      end

      def base_date
        @shared.base_date
      end

      def base_timestamp
        @shared.base_timestamp
      end

      def shared_strings
        @shared.shared_strings
      end
    end
  end
end
