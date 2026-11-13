module Roo
  class Excelx
    class Drawing
      class Base
        # extend Roo::Helpers::DefaultAttrReader
        attr_reader :value, :coordinate

        def initialize(type, value, coordinate)
          @type = type
          @value = value
          @coordinate = coordinate
        end
      end
    end
  end
end
