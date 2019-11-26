require 'uri'

module Roo
  class Spreadsheet
    class << self
      def open(path, options = {})
        path      = path.respond_to?(:path) ? path.path : path
        extension = extension_for(path, options)

        begin
          Roo::CLASS_FOR_EXTENSION.fetch(extension).new(path, options)
        rescue KeyError
          raise ArgumentError,
                "Can't detect the type of #{path} - please use the :extension option to declare its type."
        end
      end

      def extension_for(path, options)
        case (extension = options.delete(:extension))
        when ::Symbol
          options[:file_warning] = :ignore
          extension
        when ::String
          options[:file_warning] = :ignore
          extension.tr('.', '').downcase.to_sym
        else
          res = ::File.extname((path =~ /\A#{::URI::DEFAULT_PARSER.make_regexp}\z/) ? ::URI.parse(::URI.encode(path)).path : path)
          res.tr('.', '').downcase.to_sym
        end
      end
    end
  end
end
