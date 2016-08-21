module Roo
  module Tempdir
    def finalize_tempdirs(object_id)
      if @tempdirs && (dirs_to_remove = @tempdirs[object_id])
        @tempdirs[object_id] = nil
        dirs_to_remove.each do |dir|
          ::FileUtils.remove_entry(dir)
        end
      end
    end

    def make_tempdir(object, prefix, root)
      root ||= ENV['ROO_TMP']
      # folder is cleaned up in .finalize_tempdirs
      ::Dir.mktmpdir("#{Roo::TEMP_PREFIX}#{prefix}", root).tap do |tmpdir|
        @tempdirs ||= {}
        if @tempdirs[object.object_id]
          @tempdirs[object.object_id] << tmpdir
        else
          @tempdirs[object.object_id] = [tmpdir]
          ObjectSpace.define_finalizer(object, method(:finalize_tempdirs))
        end
      end
    end
  end
end
