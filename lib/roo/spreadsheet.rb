module Roo
  class Spreadsheet
    class << self
      def open(file, options = {})
        file = File === file ? file.path : file

        extension =
          if options[:extension]
            options[:file_warning] = :ignore
            ".#{options[:extension]}".gsub(/[.]+/, ".")
          else
            File.extname(URI.parse(file).path)
          end

        case extension.downcase
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
