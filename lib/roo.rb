module Roo
  def self.const_missing(const_name)
    case const_name
    else
      super
    end
  end

  autoload :Spreadsheet,  'roo/spreadsheet'
  autoload :Base,         'roo/base'

  autoload :OpenOffice,   'roo/openoffice'
  autoload :LibreOffice,  'roo/libre_office'
  autoload :Excel,        'roo/excel'
  autoload :Excelx,       'roo/excelx'
  autoload :Excel2003XML, 'roo/excel2003xml'
  autoload :Google,       'roo/google'
  autoload :CSV,          'roo/csv'
end
