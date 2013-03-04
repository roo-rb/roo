# -*- encoding: utf-8 -*-

Gem::Specification.new do |s|
  s.name = "roo"
  s.version = "1.10.3"

  s.required_rubygems_version = Gem::Requirement.new(">= 0") if s.respond_to? :required_rubygems_version=
  s.authors = ["Thomas Preymesser", "Hugh McGowan", "Ben Woosley"]
  s.date = "2013-03-04"
  s.description = "Roo can access the contents of various spreadsheet files. It can handle\n* Openoffice\n* Excel\n* Google spreadsheets\n* Excelx\n* Libreoffice\n* CSV"
  s.email = "thopre@gmail.com"
  s.executables = ["roo"]
  s.extra_rdoc_files = ["History.txt", "License.txt", "Manifest.txt", "bin/roo", "test/files/no_spreadsheet_file.txt", "website/index.txt"]
  s.files = [".gitignore", "Gemfile", "Gemfile.lock", "History.txt", "License.txt", "Manifest.txt", "README.markdown", "Rakefile", "TODO", "bin/roo", "examples/roo_soap_client.rb", "examples/roo_soap_server.rb", "examples/write_me.rb", "lib/roo.rb", "lib/roo/csv.rb", "lib/roo/excel.rb", "lib/roo/excel2003xml.rb", "lib/roo/excelx.rb", "lib/roo/generic_spreadsheet.rb", "lib/roo/google.rb", "lib/roo/openoffice.rb", "lib/roo/roo_rails_helper.rb", "lib/roo/worksheet.rb", "rm_sub_test.rb", "rm_test.rb", "roo.gemspec", "scripts/txt2html", "test/all_ss.rb", "test/files/1900_base.xls", "test/files/1904_base.xls", "test/files/Bibelbund.csv", "test/files/Bibelbund.ods", "test/files/Bibelbund.xls", "test/files/Bibelbund.xlsx", "test/files/Bibelbund.xml", "test/files/Bibelbund1.ods", "test/files/Pfand_from_windows_phone.xlsx", "test/files/bad_excel_date.xls", "test/files/bbu.ods", "test/files/bbu.xls", "test/files/bbu.xlsx", "test/files/bbu.xml", "test/files/bode-v1.ods.zip", "test/files/bode-v1.xls.zip", "test/files/boolean.ods", "test/files/boolean.xls", "test/files/boolean.xlsx", "test/files/boolean.xml", "test/files/borders.ods", "test/files/borders.xls", "test/files/borders.xlsx", "test/files/borders.xml", "test/files/bug-row-column-fixnum-float.xls", "test/files/bug-row-column-fixnum-float.xml", "test/files/comments.ods", "test/files/comments.xls", "test/files/comments.xlsx", "test/files/csvtypes.csv", "test/files/datetime.ods", "test/files/datetime.xls", "test/files/datetime.xlsx", "test/files/datetime.xml", "test/files/datetime_floatconv.xls", "test/files/datetime_floatconv.xml", "test/files/dreimalvier.ods", "test/files/emptysheets.ods", "test/files/emptysheets.xls", "test/files/emptysheets.xlsx", "test/files/emptysheets.xml", "test/files/excel2003.xml", "test/files/false_encoding.xls", "test/files/false_encoding.xml", "test/files/formula.ods", "test/files/formula.xls", "test/files/formula.xlsx", "test/files/formula.xml", "test/files/formula_parse_error.xls", "test/files/formula_parse_error.xml", "test/files/formula_string_error.xlsx", "test/files/html-escape.ods", "test/files/matrix.ods", "test/files/matrix.xls", "test/files/named_cells.ods", "test/files/named_cells.xls", "test/files/named_cells.xlsx", "test/files/no_spreadsheet_file.txt", "test/files/numbers1.csv", "test/files/numbers1.ods", "test/files/numbers1.xls", "test/files/numbers1.xlsx", "test/files/numbers1.xml", "test/files/only_one_sheet.ods", "test/files/only_one_sheet.xls", "test/files/only_one_sheet.xlsx", "test/files/only_one_sheet.xml", "test/files/paragraph.ods", "test/files/paragraph.xls", "test/files/paragraph.xlsx", "test/files/paragraph.xml", "test/files/prova.xls", "test/files/ric.ods", "test/files/simple_spreadsheet.ods", "test/files/simple_spreadsheet.xls", "test/files/simple_spreadsheet.xlsx", "test/files/simple_spreadsheet.xml", "test/files/simple_spreadsheet_from_italo.ods", "test/files/simple_spreadsheet_from_italo.xls", "test/files/simple_spreadsheet_from_italo.xml", "test/files/so_datetime.csv", "test/files/style.ods", "test/files/style.xls", "test/files/style.xlsx", "test/files/style.xml", "test/files/time-test.csv", "test/files/time-test.ods", "test/files/time-test.xls", "test/files/time-test.xlsx", "test/files/time-test.xml", "test/files/type_excel.ods", "test/files/type_excel.xlsx", "test/files/type_excelx.ods", "test/files/type_excelx.xls", "test/files/type_openoffice.xls", "test/files/type_openoffice.xlsx", "test/files/whitespace.ods", "test/files/whitespace.xls", "test/files/whitespace.xlsx", "test/files/whitespace.xml", "test/test_generic_spreadsheet.rb", "test/test_helper.rb", "test/test_roo.rb", "website/index.html", "website/index.txt", "website/javascripts/rounded_corners_lite.inc.js", "website/stylesheets/screen.css", "website/template.rhtml"]
  s.homepage = "http://roo.rubyforge.org/"
  s.rdoc_options = ["--main", "README.txt"]
  s.require_paths = ["lib"]
  s.rubyforge_project = "roo"
  s.rubygems_version = "1.8.24"
  s.summary = "Roo can access the contents of various spreadsheet files."
  s.test_files = ["test/test_generic_spreadsheet.rb", "test/test_helper.rb", "test/test_roo.rb"]

  if s.respond_to? :specification_version then
    s.specification_version = 3

    if Gem::Version.new(Gem::VERSION) >= Gem::Version.new('1.2.0') then
      s.add_runtime_dependency(%q<spreadsheet>, ["> 0.6.4"])
      s.add_runtime_dependency(%q<nokogiri>, [">= 1.4.0"])
      s.add_runtime_dependency(%q<rubyzip>, [">= 0.9.9"])
      s.add_development_dependency(%q<bones>, [">= 3.8.0"])
    else
      s.add_dependency(%q<spreadsheet>, ["> 0.6.4"])
      s.add_dependency(%q<nokogiri>, [">= 1.4.0"])
      s.add_dependency(%q<rubyzip>, [">= 0.9.9"])
      s.add_dependency(%q<bones>, [">= 3.8.0"])
    end
  else
    s.add_dependency(%q<spreadsheet>, ["> 0.6.4"])
    s.add_dependency(%q<nokogiri>, [">= 1.4.0"])
    s.add_dependency(%q<rubyzip>, [">= 0.9.9"])
    s.add_dependency(%q<bones>, [">= 3.8.0"])
  end
end
