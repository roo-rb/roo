module Roo
  class Spreadsheet
    class << self
      def open(file, options = {})
        file = File === file ? file.path : file
        case File.extname(file).downcase.match(/(.*)\?.*/)[1]
        when '.xls'
          Roo::Excel.new(file)
        when '.xlsx'
          Roo::Excelx.new(file)
        when '.ods'
          Roo::Openoffice.new(file)
        when '.xml'
          Roo::Excel2003XML.new(file)
        when ''
          Roo::Google.new(file, options)
        when '.csv'
          Roo::Csv.new(file, options)
        else
          raise ArgumentError, "Don't know how to open file #{file}"
        end
      end
    end
  end
end
