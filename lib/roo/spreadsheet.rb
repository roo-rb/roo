module Roo
  class Spreadsheet
    class << self
      def open(file)
        file = File === file ? file.path : file
        case File.extname(file).downcase
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
end
