# Tests for "Accessing Files" which includes opening and closing files.
module TestAccesingFiles
  def test_close
    with_each_spreadsheet(name: "numbers1") do |oo|
      next unless (tempdir = oo.instance_variable_get("@tmpdir"))
      oo.close
      refute File.exist?(tempdir), "Expected #{tempdir} to be cleaned up"
    end
  end

  # NOTE: Ruby 2.4.0 changed the way GC works. The last Roo object created by
  #       with_each_spreadsheet wasn't getting GC'd until after the process
  #       ended.
  #
  #       That behavior change broke this test. In order to fix it, I forked the
  #       process and passed the temp directories from the forked process in
  #       order to check if they were removed properly.
  def test_finalize
    skip if defined? JRUBY_VERSION

    read, write = IO.pipe
    pid = Process.fork do
      with_each_spreadsheet(name: "numbers1") do |oo|
        write.puts oo.instance_variable_get("@tmpdir")
      end
    end

    Process.wait(pid)
    write.close
    tempdirs = read.read.split("\n")
    read.close

    refute tempdirs.empty?
    tempdirs.each do |tempdir|
      refute File.exist?(tempdir), "Expected #{tempdir} to be cleaned up"
    end
  end

  def test_cleanup_on_error
    # NOTE: This test was occasionally failing because when it started running
    #       other tests would have already added folders to the temp directory,
    #       polluting the directory. You'd end up in a situation where there
    #       would be less folders AFTER this ran than originally started.
    #
    #       Instead, just use a custom temp directory to test the functionality.
    ENV["ROO_TMP"] = Dir.tmpdir + "/test_cleanup_on_error"
    Dir.mkdir(ENV["ROO_TMP"]) unless File.exist?(ENV["ROO_TMP"])
    expected_dir_contents = Dir.open(ENV["ROO_TMP"]).to_a
    with_each_spreadsheet(name: "non_existent_file", ignore_errors: true) {}

    assert_equal expected_dir_contents, Dir.open(ENV["ROO_TMP"]).to_a
    Dir.rmdir ENV["ROO_TMP"] if File.exist?(ENV["ROO_TMP"])
    ENV.delete "ROO_TMP"
  end

  def test_name_with_leading_slash
    xlsx = Roo::Excelx.new(File.join(TESTDIR, "/name_with_leading_slash.xlsx"))
    assert_equal 1, xlsx.sheets.count
  end
end
