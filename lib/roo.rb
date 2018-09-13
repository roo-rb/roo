# frozen_string_literal: true

require 'roo/version'
require 'roo/constants'
require 'roo/errors'
require 'roo/spreadsheet'
require 'roo/base'

module Roo
  autoload :OpenOffice,   'roo/open_office'
  autoload :LibreOffice,  'roo/libre_office'
  autoload :Excelx,       'roo/excelx'
  autoload :CSV,          'roo/csv'

  TEMP_PREFIX = 'roo_'

  CLASS_FOR_EXTENSION = {
    ods: Roo::OpenOffice,
    xlsx: Roo::Excelx,
    xlsm: Roo::Excelx,
    csv: Roo::CSV
  }

  def self.const_missing(const_name)
    case const_name
    when :Excel
      raise ROO_EXCEL_NOTICE
    when :Excel2003XML
      raise ROO_EXCELML_NOTICE
    when :Google
      raise ROO_GOOGLE_NOTICE
    else
      super
    end
  end
end
