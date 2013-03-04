module Roo

  VERSION = '1.10.3'

  class Spreadsheet
    class << self
      def open(file)
        file = File === file ? file.path : file
        case File.extname(file)
        when '.xls'
          Roo::Excel.new(file)
        when '.xlsx'
          Roo::Excelx.new(file)
        when '.ods'
          Roo::Openoffice.new(file)
        when '.xml'
          Roo::Excel2003XML.new(file)
        when ''
          Roo::Google.new(file)
        when '.csv'
          Roo::Csv.new(file)
        else
          raise ArgumentError, "Don't know how to open file #{file}"
        end
      end
    end
  end

  autoload :GenericSpreadsheet, 'roo/generic_spreadsheet'
  autoload :Openoffice,         'roo/openoffice'
  autoload :Excel,              'roo/excel'
  autoload :Excelx,             'roo/excelx'
  autoload :Google,             'roo/google'
  autoload :Csv,                'roo/csv'

  autoload :Excel2003XML,       'roo/excel2003xml'
  autoload :RooRailsHelper,     'roo/roo_rails_helper'
end
