require 'roo'

Dir.glob('test/files/*.ods').each do |fn|
  begin
    oo = Roo::OpenOffice.new fn
    print "#{File.basename(fn)} "
    puts oo.officeversion
  rescue Zip::ZipError, Errno::ENOENT => e
    # file is not a real .ods spreadsheet file
    puts e.message
  end
end
