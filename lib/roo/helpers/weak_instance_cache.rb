# frozen_string_literal: true

require "weakref"

module Roo
  module Helpers
    module WeakInstanceCache
      private

      def instance_cache(key)
        object = nil

        if instance_variable_defined?(key) && (ref = instance_variable_get(key)) && ref.weakref_alive?
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
          ObjectSpace.define_finalizer(object, instance_cache_finalizer(key))
          instance_variable_set(key, WeakRef.new(object))
        end

        object
      end

      def instance_cache_finalizer(key)
        proc do |object_id|
          if instance_variable_defined?(key) && (ref = instance_variable_get(key)) && (!ref.weakref_alive? || ref.__getobj__.object_id == object_id)
            remove_instance_variable(key)
          end
        end
      end
    end
  end
end
