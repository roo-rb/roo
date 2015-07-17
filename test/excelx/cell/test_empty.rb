require 'roo/excelx/cell/base'
require 'roo/excelx/cell/empty'

class TestRooExcelxCellEmpty < Minitest::Test
  def empty
    Roo::Excelx::Cell::Empty
  end
end
