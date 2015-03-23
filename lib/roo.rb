module Roo
  require 'roo/spreadsheet'
  require 'roo/base'
  require 'roo/open_office'
  require 'roo/libre_office'
  require 'roo/excelx'
  require 'roo/csv'

  CLASS_FOR_EXTENSION = {
    ods: Roo::OpenOffice,
    xlsx: Roo::Excelx,
    csv: Roo::CSV
  }

  def self.const_missing(const_name)
    case const_name
    when :Excel
      raise "Excel support has been extracted to roo-xls due to its dependency on the GPL'd spreadsheet gem. Install roo-xls to use Roo::Excel."
    when :Excel2003XML
      raise "Excel SpreadsheetML support has been extracted to roo-xls. Install roo-xls to use Roo::Excel2003XML."
    when :Google
      raise "Google support has been extracted to roo-google. Install roo-google to use Roo::Google."
    else
      super
    end
  end
end
