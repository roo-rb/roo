# -*- encoding: utf-8 -*-

Gem::Specification.new do |s|
  s.name = %q{roo}
  s.version = "1.2.4"

  s.required_rubygems_version = Gem::Requirement.new(">= 0") if s.respond_to? :required_rubygems_version=
  s.authors = ["Thomas Preymesser"]
  s.date = %q{2009-03-12}
  s.description = %q{roo can access the contents of OpenOffice-, Excel- or Google-Spreadsheets}
  s.email = %q{thopre@gmail.com}
  s.extra_rdoc_files = ["README.txt", "History.txt"]
  s.files = ["lib/roo", "lib/roo/excel.rb", "lib/roo/excelx.rb", "lib/roo/generic_spreadsheet.rb", "lib/roo/google.rb", "lib/roo/openoffice.rb", "lib/roo/roo_rails_helper.rb", "lib/roo/version.rb", "lib/roo.rb", "test/_ods", "test/_ods/Configurations2", "test/_ods/Configurations2/accelerator", "test/_ods/Configurations2/accelerator/current.xml", "test/_ods/Configurations2/floater", "test/_ods/Configurations2/images", "test/_ods/Configurations2/images/Bitmaps", "test/_ods/Configurations2/menubar", "test/_ods/Configurations2/popupmenu", "test/_ods/Configurations2/progressbar", "test/_ods/Configurations2/statusbar", "test/_ods/Configurations2/toolbar", "test/_ods/content.xml", "test/_ods/META-INF", "test/_ods/META-INF/manifest.xml", "test/_ods/meta.xml", "test/_ods/mimetype", "test/_ods/settings.xml", "test/_ods/style.ods", "test/_ods/styles.xml", "test/_ods/Thumbnails", "test/_ods/Thumbnails/thumbnail.png", "test/_ods_old", "test/_ods_old/Configurations2", "test/_ods_old/Configurations2/accelerator", "test/_ods_old/Configurations2/accelerator/current.xml", "test/_ods_old/Configurations2/floater", "test/_ods_old/Configurations2/images", "test/_ods_old/Configurations2/images/Bitmaps", "test/_ods_old/Configurations2/menubar", "test/_ods_old/Configurations2/popupmenu", "test/_ods_old/Configurations2/progressbar", "test/_ods_old/Configurations2/statusbar", "test/_ods_old/Configurations2/toolbar", "test/_ods_old/content.xml", "test/_ods_old/META-INF", "test/_ods_old/META-INF/manifest.xml", "test/_ods_old/meta.xml", "test/_ods_old/mimetype", "test/_ods_old/settings.xml", "test/_ods_old/style.ods", "test/_ods_old/styles.xml", "test/_ods_old/Thumbnails", "test/_ods_old/Thumbnails/thumbnail.png", "test/_xlsx", "test/_xlsx/[Content_Types].xml", "test/_xlsx/_rels", "test/_xlsx/docProps", "test/_xlsx/docProps/app.xml", "test/_xlsx/docProps/core.xml", "test/_xlsx/style.xlsx", "test/_xlsx/style.xlsx.cpgz", "test/_xlsx/xl", "test/_xlsx/xl/_rels", "test/_xlsx/xl/_rels/workbook.xml.rels", "test/_xlsx/xl/printerSettings", "test/_xlsx/xl/printerSettings/printerSettings1.bin", "test/_xlsx/xl/sharedStrings.xml", "test/_xlsx/xl/styles.xml", "test/_xlsx/xl/theme", "test/_xlsx/xl/theme/theme1.xml", "test/_xlsx/xl/workbook.xml", "test/_xlsx/xl/worksheets", "test/_xlsx/xl/worksheets/_rels", "test/_xlsx/xl/worksheets/_rels/sheet1.xml.rels", "test/_xlsx/xl/worksheets/sheet1.xml", "test/_xlsx/xl/worksheets/sheet2.xml", "test/_xlsx/xl/worksheets/sheet3.xml", "test/bbu.ods", "test/bbu.xls", "test/bbu.xlsx", "test/Bibelbund.csv", "test/Bibelbund.ods", "test/Bibelbund.xls", "test/Bibelbund.xlsx", "test/Bibelbund1.ods", "test/bode-v1.ods.zip", "test/bode-v1.xls.zip", "test/boolean.ods", "test/boolean.xls", "test/boolean.xlsx", "test/borders.ods", "test/borders.xls", "test/borders.xlsx", "test/bug-row-column-fixnum-float.xls", "test/datetime.ods", "test/datetime.xls", "test/datetime.xlsx", "test/emptysheets.ods", "test/emptysheets.xls", "test/false_encoding.xls", "test/formula.ods", "test/formula.xls", "test/formula.xlsx", "test/html-escape.ods", "test/no_spreadsheet_file.txt", "test/numbers1.csv", "test/numbers1.ods", "test/numbers1.xls", "test/numbers1.xlsx", "test/numbers1_excel.csv", "test/only_one_sheet.ods", "test/only_one_sheet.xls", "test/only_one_sheet.xlsx", "test/ric.ods", "test/simple_spreadsheet.ods", "test/simple_spreadsheet.xls", "test/simple_spreadsheet.xlsx", "test/simple_spreadsheet_from_italo.ods", "test/simple_spreadsheet_from_italo.xls", "test/style.ods", "test/style.xls", "test/style.xlsx", "test/test_helper.rb", "test/test_roo.rb", "test/time-test.csv", "test/time-test.ods", "test/time-test.xls", "test/time-test.xlsx", "README.txt", "History.txt"]
  s.has_rdoc = true
  s.homepage = %q{http://roo.rubyforge.org}
  s.rdoc_options = ["--main", "README.txt", "--inline-source", "--charset=UTF-8"]
  s.require_paths = ["lib"]
  s.rubyforge_project = %q{roo}
  s.rubygems_version = %q{1.3.1}
  s.summary = %q{roo}

  if s.respond_to? :specification_version then
    current_version = Gem::Specification::CURRENT_SPECIFICATION_VERSION
    s.specification_version = 2

    if Gem::Version.new(Gem::RubyGemsVersion) >= Gem::Version.new('1.2.0') then
      s.add_runtime_dependency(%q<spreadsheet>, [">= 0.6.3.1"])
      s.add_runtime_dependency(%q<rubyzip>, [">= 0.9.1"])
      s.add_runtime_dependency(%q<hpricot>, [">= 0.5"])
      s.add_runtime_dependency(%q<GData>, [">= 0.0.3"])
    else
      s.add_dependency(%q<spreadsheet>, [">= 0.6.3.1"])
      s.add_dependency(%q<rubyzip>, [">= 0.9.1"])
      s.add_dependency(%q<hpricot>, [">= 0.5"])
      s.add_dependency(%q<GData>, [">= 0.0.3"])
    end
  else
    s.add_dependency(%q<spreadsheet>, [">= 0.6.3.1"])
    s.add_dependency(%q<rubyzip>, [">= 0.9.1"])
    s.add_dependency(%q<hpricot>, [">= 0.5"])
    s.add_dependency(%q<GData>, [">= 0.0.3"])
  end
end
