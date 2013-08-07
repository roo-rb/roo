module Roo
  class Spreadsheet
    class << self
      def open(file, options = {})
        file = File === file ? file.path : file
        case File.extname(file).downcase
        when '.xls'
          Roo::Excel.new(file, options)
        when '.xlsx'
          Roo::Excelx.new(file, options)
        when '.ods'
          Roo::OpenOffice.new(file, options)
        when '.xml'
          Roo::Excel2003XML.new(file, options)
        when ''
          Roo::Google.new(file, options)
        when '.csv'
          Roo::CSV.new(file, options)
        else
          raise ArgumentError, "Don't know how to open file #{file}"
        end
      end
    end
  end
end
