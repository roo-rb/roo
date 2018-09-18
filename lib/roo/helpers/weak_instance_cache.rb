# frozen_string_literal: true

require "weakref"

module Roo
  module Helpers
    module WeakInstanceCache
      private

      def instance_cache(key)
        object = nil

        if (ref = instance_variable_get(key)) && ref.weakref_alive?
          begin
            object = ref.__getobj__
          rescue => e
            unless (defined?(::WeakRef::RefError) && e.is_a?(::WeakRef::RefError)) || (defined?(RefError) && e.is_a?(RefError))
              raise e
            end
          end
        end

        unless object
          object = yield
          instance_variable_set(key, WeakRef.new(object))
        end

        object
      end
    end
  end
end
