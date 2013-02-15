module Roo

  VERSION = '1.10.2'

  class Spreadsheet
    class << self
      def open(file, opts = {})
        file = File === file ? file.path : file
        
        extension = opts[:extension] ? ".#{opts[:extension]}" : File.extname(file)
        file_warning = opts[:extension] ? :ignore : :error
        
        case extension
        when '.xls'
          Roo::Excel.new(file, nil, file_warning)
        when '.xlsx'
          Roo::Excelx.new(file, nil, file_warning)
        when '.ods'
          Roo::Openoffice.new(file, nil, file_warning)
        when '.xml'
          Roo::Excel2003XML.new(file, nil, file_warning)
        when ''
          Roo::Google.new(file)
        when '.csv'
          Roo::Csv.new(file, nil, file_warning)
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
