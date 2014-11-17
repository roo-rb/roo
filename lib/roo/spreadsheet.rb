module Roo
  class Spreadsheet
    class << self
      def open(path, options = {})
        path = path.respond_to?(:path) ? path.path : path
        extension = extension_for(path, options)

        begin
          const_get(
            Roo::CLASS_FOR_EXTENSION.fetch(extension.downcase)
          ).new(path, options)
        rescue KeyError
          raise ArgumentError,
            "Can't detect the type of #{path} - please use the :extension option to declare its type."
        end
      end

      def extension_for(path, options)
        if options[:extension]
          options[:file_warning] = :ignore
          ".#{options.delete(:extension)}".gsub(/[.]+/, ".")
        else
          File.extname((path =~ URI::regexp) ? URI.parse(URI.encode(path)).path : path)
        end
      end
    end
  end
end
