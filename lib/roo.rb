module Roo

  VERSION = '1.10.3'

  autoload :Spreadsheet, 'roo/spreadsheet'

  autoload :GenericSpreadsheet, 'roo/generic_spreadsheet'
  autoload :Openoffice,         'roo/openoffice'
  autoload :Excel,              'roo/excel'
  autoload :Excelx,             'roo/excelx'
  autoload :Google,             'roo/google'
  autoload :Csv,                'roo/csv'

  autoload :Excel2003XML,       'roo/excel2003xml'
  autoload :RooRailsHelper,     'roo/roo_rails_helper'
end
