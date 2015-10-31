module Roo
  # Public: methods for opening and closing spreadsheet files.
  module FileUtils
    TEMPDIR_PREFIX = 'roo_'.freeze

    def close
      return nil unless @tmpdirs
      @tmpdirs.each { |dir| ::FileUtils.remove_entry(dir) }

      nil
    end

    private

    def is_stream?(filename_or_stream)
      filename_or_stream.respond_to?(:seek)
    end

    # move random Roo::Base functions to Utils
    def make_tmpdir(prefix = nil, root = nil, &block)
      prefix ||= TEMPDIR_PREFIX

      ::Dir.mktmpdir(prefix, root, &block).tap do |result|
        block_given? || track_tmpdir!(result)
      end
    end

    def local_filename(filename, tmpdir, packed = nil, stream = false)
      # FIXME: Remove nil check if Roo.open_new validates path
      fail IOError, 'filename cannot be nil' unless filename
      return if stream || is_stream?(filename)

      filename = download_uri(filename, tmpdir) if uri?(filename)
      filename = unzip(filename, tmpdir) if packed == :zip

      fail IOError, "#{filename} does not exist" unless File.exist?(filename)

      filename
    end

    def uri?(filename)
      filename.start_with?('http://', 'https://')
    end

    def download_uri(uri, tmpdir)
      require 'open-uri'
      tempfilename = File.join(tmpdir, File.basename(uri))
      begin
        File.open(tempfilename, 'wb') do |file|
          open(uri, 'User-Agent' => "Ruby/#{RUBY_VERSION}") do |net|
            file.write(net.read)
          end
        end
      rescue OpenURI::HTTPError
        raise "could not open #{uri}"
      end

      tempfilename
    end

    def track_tmpdir!(tmpdir)
      (@tmpdirs ||= []) << tmpdir
    end

    def unzip(filename, tmpdir)
      require 'zip/filesystem'

      Zip::File.open(filename) do |zip|
        process_zipfile_packed(zip, tmpdir)
      end
    end

    def process_zipfile_packed(zip, tmpdir, path = '')
      if zip.file.file? path
        # extract and return filename
        File.open(File.join(tmpdir, path), 'wb') do |file|
          file.write(zip.read(path))
        end
        File.join(tmpdir, path)
      else
        ret = nil
        path += '/' unless path.empty?
        zip.dir.foreach(path) do |filename|
          ret = process_zipfile_packed(zip, tmpdir, path + filename)
        end
        ret
      end
    end
  end
end
