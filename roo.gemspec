# -*- encoding: utf-8 -*-

Gem::Specification.new do |s|
  s.name = %q{roo}
  s.version = "1.3.0"

  s.required_rubygems_version = Gem::Requirement.new(">= 0") if s.respond_to? :required_rubygems_version=
  s.authors = ["Thomas Preymesser"]
  s.date = %q{2009-04-20}
  s.description = %q{roo can access the contents of OpenOffice-, Excel- or Google-Spreadsheets}
  s.email = %q{thopre@gmail.com}
  s.extra_rdoc_files = ["README.markdown", "History.txt"]
  s.files = ["lib/roo", "lib/roo/excel.rb", "lib/roo/excelx.rb", "lib/roo/generic_spreadsheet.rb", "lib/roo/google.rb", "lib/roo/openoffice.rb", "lib/roo/roo_rails_helper.rb", "lib/roo/version.rb", "lib/roo.rb", "test/bbu.ods", "test/bbu.xls", "test/bbu.xlsx", "test/Bibelbund.csv", "test/Bibelbund.ods", "test/Bibelbund.xls", "test/Bibelbund.xlsx", "test/Bibelbund1.ods", "test/bode-v1.ods.zip", "test/bode-v1.xls.zip", "test/boolean.ods", "test/boolean.xls", "test/boolean.xlsx", "test/borders.ods", "test/borders.xls", "test/borders.xlsx", "test/bug-row-column-fixnum-float.xls", "test/datetime.ods", "test/datetime.xls", "test/datetime.xlsx", "test/datetime_floatconv.xls", "test/emptysheets.ods", "test/emptysheets.xls", "test/false_encoding.xls", "test/formula.ods", "test/formula.xls", "test/formula.xlsx", "test/html-escape.ods", "test/no_spreadsheet_file.txt", "test/numbers1.csv", "test/numbers1.ods", "test/numbers1.xls", "test/numbers1.xlsx", "test/numbers1_excel.csv", "test/only_one_sheet.ods", "test/only_one_sheet.xls", "test/only_one_sheet.xlsx", "test/paragraph.ods", "test/paragraph.xls", "test/paragraph.xlsx", "test/ric.ods", "test/simple_spreadsheet.ods", "test/simple_spreadsheet.xls", "test/simple_spreadsheet.xlsx", "test/simple_spreadsheet_from_italo.ods", "test/simple_spreadsheet_from_italo.xls", "test/style.ods", "test/style.xls", "test/style.xlsx", "test/test_helper.rb", "test/test_roo.rb", "test/time-test.csv", "test/time-test.ods", "test/time-test.xls", "test/time-test.xlsx", "README.markdown", "History.txt"]
  s.has_rdoc = true
  s.homepage = %q{http://roo.rubyforge.org}
  s.rdoc_options = ["--main", "README.markdown", "--inline-source", "--charset=UTF-8"]
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
