module Roo
  module Tempdir
    def finalize_tempdirs(object_id)
      if @tempdirs && (dirs_to_remove = @tempdirs[object_id])
        @tempdirs.delete(object_id)
        dirs_to_remove.each do |dir|
          # Pass force=true to avoid an exception (and thus warnings in Ruby 3.1) if dir has
          # already been removed. This can occur when the finalizer is called both in a forked
          # child process and in the parent.
          ::FileUtils.remove_entry(dir, true)
        end
      end
    end

    def make_tempdir(object, prefix, root)
      root ||= ENV["ROO_TMP"]
      # NOTE: This folder is cleaned up by finalize_tempdirs.
      ::Dir.mktmpdir("#{Roo::TEMP_PREFIX}#{prefix}", root).tap do |tmpdir|
        @tempdirs ||= Hash.new { |h, k| h[k] = [] }
        @tempdirs[object.object_id] << tmpdir
      end
    end
  end
end
