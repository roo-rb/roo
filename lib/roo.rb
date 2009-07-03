module Roo
  class Spreadsheet
    class << self
      def open(file)
        case File.extname(file)
        when '.xls'
          Excel.new(file)
        when '.xlsx'
          Excelx.new(file)
        when '.ods'
          Openoffice.new(file)
        when ''
          Google.new(file)
        else
          raise ArgumentError, "Don't know how to open file #{file}"
        end      
      end
    end
  end
end

require 'roo/version'
# require 'roo/spreadsheetparser' TODO:
require 'roo/generic_spreadsheet' 
require 'roo/openoffice'
require 'roo/excel'
require 'roo/excelx'
require 'roo/google'
require 'roo/roo_rails_helper'
