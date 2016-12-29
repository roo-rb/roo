module Roo
  module Tempdir
    def finalize_tempdirs(object_id)
      if @tempdirs && (dirs_to_remove = @tempdirs[object_id])
        @tempdirs.delete(object_id)
        dirs_to_remove.each do |dir|
          ::FileUtils.remove_entry(dir)
        end
      end
    end

    def make_tempdir(object, prefix, root)
      root ||= ENV["ROO_TMP"]
      # folder is cleaned up in .finalize_tempdirs
      ::Dir.mktmpdir("#{Roo::TEMP_PREFIX}#{prefix}", root).tap do |tmpdir|
        @tempdirs ||= Hash.new { |h, k| h[k] = [] }

        if @tempdirs[object.object_id].empty?
          ObjectSpace.define_finalizer(object, method(:finalize_tempdirs))
        end

        @tempdirs[object.object_id] << tmpdir
      end
    end
  end
end
