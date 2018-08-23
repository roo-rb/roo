module Roo
  module Tempdir
    def finalize_tempdirs(object_id)
      if @tempdirs && (dirs_to_remove = @tempdirs[object_id])
        @tempdirs.delete(object_id)

        # Ignore the directory list since it seems this method gets called too
        # prematurely, related to the object being garbage collected in a
        # different process (sidekiq/puma for example)

        cmds = [
          # Delete roo temporary directories over 15 minutes old.
          "find #{ENV["ROO_TMP"]||'/tmp'} -type d -name '#{Roo::TEMP_PREFIX}*' -mmin +15 -exec rm -r {} \\;",

          # Uploaded files (I think from roo)
          "find #{ENV["ROO_TMP"]||'/tmp'} -name 'RackMultipart*' -mmin +15 -exec rm -r {} \\;"
        ]

        cmds.each do |cmd|
          system(cmd)
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
