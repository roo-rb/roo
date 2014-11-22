require 'roo/excelx/extractor'

module Roo
  class Excelx::SheetDoc < Excelx::Extractor
    def initialize(path, relationships, styles, shared_strings, workbook)
      super(path)
      @relationships = relationships
      @styles = styles
      @shared_strings = shared_strings
      @workbook = workbook
    end

    def cells(relationships)
      @cells ||=
        Hash[doc.xpath("/worksheet/sheetData/row/c").map do |cell_xml|
          key = Roo::Base.ref_to_key(cell_xml['r'])
          [key, cell_from_xml(cell_xml, hyperlinks(relationships)[key])]
        end]
    end

    def hyperlinks(relationships)
      @hyperlinks ||=
        Hash[doc.xpath("/worksheet/hyperlinks/hyperlink").map do |hyperlink|
          if hyperlink.attribute('id') && relationship = relationships[hyperlink.attribute('id').text]
            [Roo::Base.ref_to_key(hyperlink.attributes['ref'].to_s), relationship.attribute('Target').text]
          end
        end.compact]
    end

    private

    def cell_from_xml(cell_xml, hyperlink)
      style = cell_xml['s'].to_i   # should be here
      # c: <c r="A5" s="2">
      # <v>22606</v>
      # </c>, format: , tmp_type: float
      value_type =
        case cell_xml['t']
        when 's'
          :shared
        when 'b'
          :boolean
        # 2011-02-25 BEGIN
        when 'str'
          :string
        # 2011-02-25 END
        # 2011-09-15 BEGIN
        when 'inlineStr'
          :inlinestr
        # 2011-09-15 END
        else
          format = @styles.style_format(style)
          Excelx::Format.to_type(format)
        end
      formula = nil
      cell_xml.children.each do |cell|
        case cell.name
        when 'is'
          cell.children.each do |inline_str|
            if inline_str.name == 't'
              return Excelx::Cell.new(inline_str.content,:string,formula,:string,inline_str.content,style, hyperlink, @workbook.base_date)
            end
          end
        when 'f'
          formula = cell.content
        when 'v'
          if [:time, :datetime].include?(value_type) && cell.content.to_f >= 1.0
            value_type =
              if (cell.content.to_f - cell.content.to_f.floor).abs > 0.000001
                :datetime
              else
                :date
              end
          end
          excelx_type = [:numeric_or_formula,format.to_s]
          value =
            case value_type
            when :shared
              value_type = :string
              excelx_type = :string
              @shared_strings[cell.content.to_i]
            when :boolean
              (cell.content.to_i == 1 ? 'TRUE' : 'FALSE')
            when :date, :time, :datetime
              cell.content
            when :formula
              cell.content.to_f
            when :string
              excelx_type = :string
              cell.content
            else
              value_type = :float
              cell.content
            end
          return Excelx::Cell.new(value,value_type,formula,excelx_type,cell.content,style, hyperlink, @workbook.base_date)
        end
      end
      nil
    end

=begin
Datei xl/comments1.xml
  <?xml version="1.0" encoding="UTF-8" standalone="yes" ?>
  <comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <authors>
      <author />
    </authors>
    <commentList>
      <comment ref="B4" authorId="0">
        <text>
          <r>
            <rPr>
              <sz val="10" />
              <rFont val="Arial" />
              <family val="2" />
            </rPr>
            <t>Kommentar fuer B4</t>
          </r>
        </text>
      </comment>
      <comment ref="B5" authorId="0">
        <text>
          <r>
            <rPr>
            <sz val="10" />
            <rFont val="Arial" />
            <family val="2" />
          </rPr>
          <t>Kommentar fuer B5</t>
        </r>
      </text>
    </comment>
  </commentList>
  </comments>
=end
=begin
    if @comments_doc[self.sheets.index(sheet)]
      read_comments(sheet)
    end
=end
  end
end
