module Roo

  VERSION = File.read(File.join(__dir__, "../VERSION")) rescue "0.0.0-unknown"

  def self.const_missing(const_name)
    case const_name
    when :Libreoffice
      warn "`Roo::Libreoffice` has been deprecated. Use `Roo::LibreOffice` instead."
      LibreOffice
    when :Openoffice
      warn "`Roo::Openoffice` has been deprecated. Use `Roo::OpenOffice` instead."
      OpenOffice
    when :Csv
      warn "`Roo::Csv` has been deprecated. Use `Roo::CSV` instead."
      CSV
    when :GenericSpreadsheet
      warn "`Roo::GenericSpreadsheet` has been deprecated. Use `Roo::Base` instead."
      Base
    else
      super
    end
  end

  autoload :Spreadsheet,  'roo/spreadsheet'
  autoload :Base,         'roo/base'

  autoload :OpenOffice,   'roo/openoffice'
  autoload :LibreOffice,  'roo/openoffice'
  autoload :Excel,        'roo/excel'
  autoload :Excelx,       'roo/excelx'
  autoload :Excel2003XML, 'roo/excel2003xml'
  autoload :Google,       'roo/google'
  autoload :CSV,          'roo/csv'
end
