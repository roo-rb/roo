module Roo
  class Spreadsheet
    class << self
      def open(file)
        file = String === file ? file : file.path
        begin
          case File.extname(file)
          when '.xls'
            Excel.new(file)
          when '.xlsx'
            Excelx.new(file)
          when '.ods'
            Openoffice.new(file)
          when '.xml'
            Excel2003XML.new(file)
          when ''
            Google.new(file)
          else
            raise ArgumentError, "Don't know how to open file #{file}"
          end
        rescue Ole::Storage::FormatError
          Roo::HTML.new(file)
        end
      end
    end
  end

  # autoload :Spreadsheetparser, 'roo/spreadsheetparser' TODO:
  autoload :GenericSpreadsheet, 'roo/generic_spreadsheet'
  autoload :Openoffice,         'roo/openoffice'
  autoload :Excel,              'roo/excel'
  autoload :Excelx,             'roo/excelx'
  autoload :Google,             'roo/google'
  autoload :Excel2003XML,       'roo/excel2003xml'
  autoload :RooRailsHelper,     'roo/roo_rails_helper'
  autoload :Worksheet,          'roo/worksheet'
  autoload :HTML, 'roo/html'
end

require 'roo/version'
