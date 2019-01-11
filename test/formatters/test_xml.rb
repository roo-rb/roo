require "test_helper"

class TestRooFormatterXML < Minitest::Test
  def test_to_xml
    expected_sheet_count = 5
    options = { name: "numbers1", encoding: "utf8" }
    with_each_spreadsheet(options) do |workbook|
      skip if defined? JRUBY_VERSION
      workbook.to_xml
      sheetname = workbook.sheets.first
      doc = Nokogiri::XML(workbook.to_xml)

      assert_equal expected_sheet_count, doc.xpath("//spreadsheet/sheet").count

      doc.xpath("//spreadsheet/sheet").each do |xml_sheet|
        all_cells = init_all_cells(workbook, sheetname)
        cells = xml_sheet.children.reject(&:text?)

        assert_equal sheetname, xml_sheet["name"]
        assert_equal all_cells.size, cells.size

        cells.each_with_index do |cell, i|
          expected = [
            all_cells[i][:row],
            all_cells[i][:column],
            all_cells[i][:content],
            all_cells[i][:type],
          ]
          result = [
            cell["row"],
            cell["column"],
            cell.text,
            cell["type"],
          ]

          assert_equal expected, result
        end # end of sheet
        sheetname = workbook.sheets[workbook.sheets.index(sheetname) + 1]
      end
    end
  end

  def test_bug_to_xml_with_empty_sheets
    with_each_spreadsheet(name: "emptysheets", format: [:openoffice, :excelx]) do |workbook|
      workbook.sheets.each do |sheet|
        assert_nil workbook.first_row, "first_row not nil in sheet #{sheet}"
        assert_nil workbook.last_row, "last_row not nil in sheet #{sheet}"
        assert_nil workbook.first_column, "first_column not nil in sheet #{sheet}"
        assert_nil workbook.last_column, "last_column not nil in sheet #{sheet}"
        assert_nil workbook.first_row(sheet), "first_row not nil in sheet #{sheet}"
        assert_nil workbook.last_row(sheet), "last_row not nil in sheet #{sheet}"
        assert_nil workbook.first_column(sheet), "first_column not nil in sheet #{sheet}"
        assert_nil workbook.last_column(sheet), "last_column not nil in sheet #{sheet}"
      end
      workbook.to_xml
    end
  end

  # Erstellt eine Liste aller Zellen im Spreadsheet. Dies ist nötig, weil ein einfacher
  # Textvergleich des XML-Outputs nicht funktioniert, da xml-builder die Attribute
  # nicht immer in der gleichen Reihenfolge erzeugt.
  def init_all_cells(workbook, sheet)
    all = []
    workbook.first_row(sheet).upto(workbook.last_row(sheet)) do |row|
      workbook.first_column(sheet).upto(workbook.last_column(sheet)) do |col|
        next if workbook.empty?(row, col, sheet)

        all << {
          row: row.to_s,
          column: col.to_s,
          content: workbook.cell(row, col, sheet).to_s,
          type: workbook.celltype(row, col, sheet).to_s,
        }
      end
    end
    all
  end
end
