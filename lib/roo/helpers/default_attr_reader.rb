# frozen_string_literal: true

module Roo
  module Helpers
    module DefaultAttrReader
      def attr_reader_with_default(attr_hash)
        attr_hash.each do |attr_name, default_value|
          instance_variable = :"@#{attr_name}"
          define_method attr_name do
            if instance_variable_defined? instance_variable
              instance_variable_get instance_variable
            else
              default_value
            end
          end
        end
      end
    end
  end
end
