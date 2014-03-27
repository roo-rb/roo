module Roo
  class Spreadsheet
    class << self
      def open(path, options = {})
        path = path.respond_to?(:path) ? path.path : path

        extension =
          if options[:extension]
            options[:file_warning] = :ignore
            ".#{options.delete(:extension)}".gsub(/[.]+/, ".")
          else
            File.extname((path =~ URI::regexp) ? URI.parse(path).path : path)
          end

        case extension.downcase
        when '.xls'
          Roo::Excel.new(path, options)
        when '.xlsx'
          Roo::Excelx.new(path, options)
        when '.ods'
          Roo::OpenOffice.new(path, options)
        when '.xml'
          Roo::Excel2003XML.new(path, options)
        when ''
          Roo::Google.new(path, options)
        when '.csv'
          Roo::CSV.new(path, options)
        else
          raise ArgumentError,
            "Can't detect the type of #{path} - please use the :extension option to declare its type."
        end
      end
    end
  end
end
