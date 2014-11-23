Roo::ZipFile =
  begin
    require 'zip/zipfilesystem'
    Zip::ZipFile
  rescue LoadError
    # For rubyzip >= 1.0.0
    require 'zip/filesystem'
    Zip::File
  end
