module Roo
  class Spreadsheet
    class << self
      def open(file, options = {})
        file = file.respond_to?(:path) ? file.path : file

        extension =
          if options[:extension]
            options[:file_warning] = :ignore
            ".#{options.delete(:extension)}".gsub(/[.]+/, ".")
          else
            File.extname(URI.decode(URI.parse(URI.encode(file)).path))
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
          raise ArgumentError,
            "Can't detect the type of #{file} - please use the :extension option to declare its type."
        end
      end
    end
  end
end
