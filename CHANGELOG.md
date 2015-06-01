## [2.0.1] - 2015-06-01
### Added
- Return an enumerator when calling '#each' without a block [#219](https://github.com/roo-rb/roo/pull/219)
- Added Roo::Base#close to delete any temp directories[#211](https://github.com/roo-rb/roo/pull/211)
- Offset option for excelx #each_row. [#214](https://github.com/roo-rb/roo/pull/214)
- Allow Roo::Excelx to open streams [#209](https://github.com/roo-rb/roo/pull/209)

### Fixed
- Use gsub instead of tr for double quote escaping [#212](https://github.com/roo-rb/roo/pull/212),  [#212-patch](https://github.com/roo-rb/roo/commit/fcc9a015868ebf9d42cbba5b6cfdaa58b81ecc01)
- Fixed Changelog links and release data. [#204](https://github.com/roo-rb/roo/pull/204), [#206](https://github.com/roo-rb/roo/pull/206)
- Allow Pathnames to be used when opening files. [#207](https://github.com/roo-rb/roo/pull/207)

### Removed
- Removed the scripts folder. [#213](https://github.com/roo-rb/roo/pull/213)

## [2.0.0] - 2015-04-24
### Added
- Added optional support for hidden sheets in Excelx and LibreOffice files [#177](https://github.com/roo-rb/roo/pull/177)
- Roo::OpenOffice can be used to open encrypted workbooks. [#157](https://github.com/roo-rb/roo/pull/157)
- Added streaming for parsing of large Excelx Sheets. [#69](https://github.com/roo-rb/roo/pull/69)
- Added Roo::Base#first_last_row_col_for_sheet [a0dd800](https://github.com/roo-rb/roo/commit/a0dd800d5cf0de052583afa91bf82f8802ede9a0)
- Added Roo::Base#collect_last_row_col_for_sheet [a0dd800](https://github.com/roo-rb/roo/commit/a0dd800d5cf0de052583afa91bf82f8802ede9a0)
- Added Roo::Base::MAX_ROW_COL, Roo::Base::MIN_ROW_COL [a0dd800](https://github.com/roo-rb/roo/commit/a0dd800d5cf0de052583afa91bf82f8802ede9a0)
- Extract Roo::Font to replace equivalent uses in Excelx and OpenOffice. [23e19de](https://github.com/roo-rb/roo/commit/23e19de1ccc64b2b02a80090ff6666008a29c43b)
- Roo::Utils [3169a0e](https://github.com/roo-rb/roo/commit/3169a0e803ce742d2cbf9be834d27a5098a68638)
- Roo::ExcelxComments [0a43341](https://github.com/roo-rb/roo/commit/0a433413210b5559dc92d743a72a1f38ee775f5f)
[0a43341](https://github.com/roo-rb/roo/commit/0a433413210b5559dc92d743a72a1f38ee775f5f)
- Roo::Excelx::Relationships [0a43341](https://github.com/roo-rb/roo/commit/0a433413210b5559dc92d743a72a1f38ee775f5f)
- Roo::Excelx::SheetDoc [0a43341](https://github.com/roo-rb/roo/commit/0a433413210b5559dc92d743a72a1f38ee775f5f)
[c2bb7b8](https://github.com/roo-rb/roo/commit/c2bb7b8614f4ff1dff6b7bdbda0ded125ae549c7)
+- Roo::Excelx::Styles [c2bb7b8](https://github.com/roo-rb/roo/commit/c2bb7b8614f4ff1dff6b7bdbda0ded125ae549c7)
+- Roo::Excelx::Workbook [c2bb7b8](https://github.com/roo-rb/roo/commit/c2bb7b8614f4ff1dff6b7bdbda0ded125ae549c7)
- Switch from Spreadsheet::Link to Roo::Link [ee67321](https://github.com/roo-rb/roo/commit/ee6732144f3616631d19ade0c5490e1678231ce2)
- Roo::Base#to_csv: Added separator parameter  (defaults to ",") [#102](https://github.com/roo-rb/roo/pull/102)
- Added development development gems [#104](https://github.com/roo-rb/roo/pull/104)

### Changed
- Reduced size of published gem. [#194](https://github.com/roo-rb/roo/pull/194)
- Stream the reading of the dimensions [#192](https://github.com/roo-rb/roo/pull/192)
- Return `nil` when a querying a cell that doesn't exist (instead of a NoMethodError) [#192](https://github.com/roo-rb/roo/pull/192), [#165](https://github.com/roo-rb/roo/pull/165)
- Roo::OpenOffice#formula? now returns a `Boolean` instead of a `String` or `nil` [#191](https://github.com/roo-rb/roo/pull/191)
- Added a less verbose Roo::Base#inspect. It no longer returns the entire object. [#188](https://github.com/roo-rb/roo/pull/188), [#186](https://github.com/roo-rb/roo/pull/186)
- Memoize Roo::Utils.split_coordinate [#180](https://github.com/roo-rb/roo/pull/180)
- Roo::Base: use regular expressions for extracting headers [#173](https://github.com/roo-rb/roo/pull/173)
- Roo::Base: memoized `first_row`/`last_row` `first_column`/`last_column` and changed the default value of the `sheet` argument from `nil` to `default_sheet` [a0dd800](https://github.com/roo-rb/roo/commit/a0dd800d5cf0de052583afa91bf82f8802ede9a0)
- Roo::Base: changed the order of arguments for `to_csv` to (filename = nil, separator = ',', sheet = default_sheet) from (filename=nil,sheet=nil) [1e82a21](https://github.com/roo-rb/roo/commit/1e82a218087ba34379ae7312214911b104333e2c)
- In OpenOffice / LibreOffice, load the content xml lazily. Leave the tmpdir open so that reading may take place after initialize. The OS will be responsible for cleaning it up. [adb204b](https://github.com/roo-rb/roo/commit/a74157adb204bc93d289c5708e8e79e143d09037)
- Lazily initialize @default_sheet, to avoid reading the sheets earlier than necessary. Use the #default_sheet accessor instead. [704e3dc](https://github.com/roo-rb/roo/commit/704e3dca1692d84ac4877f04a7e46238772d423b)
- Roo::Base#default_sheet is no longer an attr_reader [704e3dc](https://github.com/roo-rb/roo/commit/704e3dca1692d84ac4877f04a7e46238772d423b)
- In Excelx, load styles, shared strings and the workbook lazily. Leave the tmpdir open so that reading may take place after initialize. The OS will be responsible for cleaning it up. [a973237](https://github.com/roo-rb/roo/commit/a9732372f531e435a3330d8ab5bd44ce2cb57b0b), [4834e20c](https://github.com/roo-rb/roo/commit/4834e20c6c4d2086414c43f8b0cc2d1413b45a61), [e49a1da](https://github.com/roo-rb/roo/commit/e49a1da22946918992873e8cd5bacc15ea2c73c4)
- Change the tmpdir prefix from oo_ to roo_ [102d5fc](https://github.com/roo-rb/roo/commit/102d5fce30b46e928807bc60f607f81956ed898b)
- Accept the tmpdir_root option in Roo::Excelx [0e325b6](https://github.com/roo-rb/roo/commit/0e325b68f199ff278b26bd621371ed42fa999f24)
- Refactored Excelx#comment? [0fb90ec](https://github.com/roo-rb/roo/commit/0fb90ecf6a8f422ef16a7105a1d2c42d611556c3)
- Refactored Roo::Base#find, #find_by_row, #find_by_conditions. [1ccedab](https://github.com/roo-rb/roo/commit/1ccedab3fb656f4614f0a85e9b0a286ad83f5c1e)
- Extended Roo::Spreadsheet.open so that it accepts Tempfiles and other arguments responding to `path`. Note they require an :extension option to be declared, as the tempfile mangles the extension. [#84](https://github.com/roo-rb/roo/pull/84).

### Fixed
- Process sheets from Numbers 3.1 xlsx files in the right order. [#196](https://github.com/roo-rb/roo/pull/196), [#181](https://github.com/roo-rb/roo/pull/181), [#114](https://github.com/roo-rb/roo/pull/114)
- Fixed comments for xlsx files exported from Google [#197](https://github.com/roo-rb/roo/pull/197)
- Fixed Roo::Excelx#celltype to return :link when appropriate.
- Fixed type coercion of ids. [#192](https://github.com/roo-rb/roo/pull/192)
- Clean option only removes spaces and control characters instead of removing all characters outside of the ASCII range. [#176](https://github.com/roo-rb/roo/pull/176)
- Fixed parse method with `clean` option [#184](https://github.com/roo-rb/roo/pull/184)
- Fixed some memory issues.
- Fixed Roo::Utils.number_to_letter [#180](https://github.com/roo-rb/roo/pull/180)
- Fixed merged cells return value. Instead of only the top-left cell returning a value, all merged cells return that value instead of returning nil. [#171](https://github.com/roo-rb/roo/pull/171)
- Handle headers with brackets [#162](https://github.com/roo-rb/roo/pull/162)
- Roo::Base#sheet method was not returning the sheet specified when using either an index or name [#160](https://github.com/roo-rb/roo/pull/160)
- Properly process paths with spaces. [#142](https://github.com/roo-rb/roo/pull/142), [#121](https://github.com/roo-rb/roo/pull/121), [#94](https://github.com/roo-rb/roo/pull/94),  [4e7d7d1](https://github.com/roo-rb/roo/commit/4e7d7d18d37654b0c73b229f31ea0d305c7e90ff)
- Disambiguate #open call in Excelx#extract_file. [#125](https://github.com/roo-rb/roo/pull/125)
- Fixed that #parse-ing with a hash of columns not in the document would fail mysteriously. [#129](https://github.com/roo-rb/roo/pull/129)
- Fixed Excelx issue when reading hyperlinks [#123](https://github.com/roo-rb/roo/pull/123)
- Fixed invalid test case [#124](https://github.com/roo-rb/roo/pull/124)
- Fixed error in test helper file_diff [56e2e61](https://github.com/roo-rb/roo/commit/56e2e61d1ad9185d8ab0d4af4b32928f07fdaad0)
- Stopped `inspect` from being called recursively. [#115](https://github.com/roo-rb/roo/pull/115)
- Fixes for Excelx Datetime cells. [#104](https://github.com/roo-rb/roo/pull/104), [#120](https://github.com/roo-rb/roo/pull/120)
- Prevent ArgumentError when using `find` [#100](https://github.com/roo-rb/roo/pull/100)
- Export to_csv converts link cells to url [#93](https://github.com/roo-rb/roo/pull/93), [#108](https://github.com/roo-rb/roo/pull/108)

### Removed
- Roo::Excel - Extracted to roo-xls gem. [a7edbec](https://github.com/roo-rb/roo/commit/a7edbec2eb44344611f82cff89a82dac31ec0d79)
- Roo::Excel2003XML - Extracted to roo-xls gem. [a7edbec](https://github.com/roo-rb/roo/commit/a7edbec2eb44344611f82cff89a82dac31ec0d79)
- Roo::Google - Extracted to roo-google gem. [a7edbec](https://github.com/roo-rb/roo/commit/a7edbec2eb44344611f82cff89a82dac31ec0d79)
- Roo::OpenOffice::Font - Refactored into Roo::Font
- Removed Roo::OpenOffice.extract_content [a74157a](https://github.com/roo-rb/roo/commit/a74157adb204bc93d289c5708e8e79e143d09037)
- Removed OpenOffice.process_zipfile [835368e](https://github.com/roo-rb/roo/commit/835368e1d29c1530f00bf9caa07704b17370e38f)
- Roo::OpenOffice#comment?
- Roo::ZipFile - Removed the Roo::ZipFile abstraction. Roo now depends on rubyzip 1.0.0+ [d466950](https://github.com/roo-rb/roo/commit/d4669503b5b80c1d30f035177a2b0e4b56fc49ce)
- SpreadSheet::Worksheet - Extracted to roo-xls gem. [a7edbec](https://github.com/roo-rb/roo/commit/a7edbec2eb44344611f82cff89a82dac31ec0d79)
- Spreadsheet - Extracted to roo-xls gem. [a7edbec](https://github.com/roo-rb/roo/commit/a7edbec2eb44344611f82cff89a82dac31ec0d79)

## [1.13.2] - 2013-12-23
### Fixed
- Fix that Excelx link-cells would blow up if the value wasn't a string. Due to the way Spreadsheet::Link is implemented the link text must be treated as a string. #92

## [1.13.1] - 2013-12-23
### Fixed
- Fix that Excelx creation could blow up due to nil rels files. #90

## [1.13.0] - 2013-12-05
### Changed / Added
- Support extracting link data from Excel and Excelx spreadsheets,
    via Excel#read_cell() and Excelx#hyperlink(?). #47
- Support setting the Excel Spreadsheet mode via the :mode option. #88
- Support Spreadsheet.open with a declared :extension that includes a leading '.'. #73
- Enable file type detection for URI's with parameters / anchors. #51

### Fixed
- Fix that CSV#each_row could overwrite the filename when run against a uri. #77
- Fix that #to_matrix wasn't respecting the sheet argument. #87

## [1.12.2] - 2013-09-11
### Changed / Added
- Support rubyzip >= 1.0.0. #65
- Fix typo in deprecation notices. #63

## [1.12.1] - 2013-08-18
### Changed / Added
- Support :boolean fields for CSV export via #cell_to_csv. #59

### Fixed
- Fix that Excelx would error on files with gaps in the numbering of their
  internal sheet#.xml files. #58
- Fix that Base#info to preserve the original value of #default_sheet. #44

## [1.12.0] - 2013-08-18
### Deprecated
- Rename Openoffice -> OpenOffice, Libreoffice -> LibreOffice, Csv -> CSV, and redirect the old names to the new constants
- Enable Roo::Excel, Excel2003XML, Excelx, OpenOffice to accept an options hash, and deprecate the old method argument based approach to supplying them options
- Roo's roo_rails_helper, aka the `spreadsheet` html-generating view method is currently deprecated with no replacement. If you find it helpful, tell http://github.com/Empact or extract it yourself.

### Changed / Added
- Add Roo::Excelx#load_xml so that people can customize to their data, e.g. #23
- Enable passing csv_options to Roo::CSV, which are passed through to the underlying CSV call.
- Enable passing options through from Roo::Spreadsheet to any Roo type.
- Enable passing an :extension option to Roo::Spreadsheet.new, which will override the extension detected on in the path #15
- Switch from google-spreadsheet-ruby to google_drive for Roo::Google access #40
- Make all the classes consistent in that #read_cells is only effective if the sheet has not been read.
- Roo::Google supports login via oauth :access_token. #61
- Roo::Excel now exposes its Spreadsheet workbook via #workbook
- Pull #load_xml down into Roo::Base, and use it in Excel2003XML and OpenOffice.

### Changed
- #formula? now returns truthy or falsey, rather than true/false.
- Base#longest_sheet was moved to Excel, as it only worked under Excel

### Fixed
- Fix that Roo::CSV#parse(headers: true) would blow up. #37

## [1.11.2] - 2013-04-10

### Fixed
- Fix that Roo::Spreadsheet.open wasn't tolerant to case differences.
- Fix that Roo::Excel2003XML loading was broken #27
- Enable loading Roo::Csv files from uris, just as other file types #31
- Fix that Excelx "m/d/yy h:mm" was improperly being interpreted as date rather
    than datetime #29

## [1.11.1] - 2013-03-18
### Fixed
- Exclude test/log/roo.log test log file from the gemspec in order to avoid a
    rubygems warning: #26

## [1.11.0] - 2013-03-14
### Changed / Added
- Support ruby 2.0.0 by replacing Iconv with String#encode #19
- Excelx: Loosen the format detection rules such that more are
    successfully detected #20
- Delete the roo binary, which was useless and not declared in the gemspec

### Changed
- Drop support for ruby 1.8.x or lower. Required in order to easily support 2.0.0.

## [1.10.3] - 2013-03-03
### Fixed
- Support both nokogiri 1.5.5 and 1.5.6 (Karsten Richter) #18

### Changed / Added
- Relax our nokogiri dependency back to 1.4.0, as we have no particular reason
    to require a newer version.

## [1.10.2] - 2013-02-03
### Fixed
- Support opening URIs with query strings https://github.com/Empact/roo/commit/abf94bdb59cabc16d4f7764025e88e3661983525
- Support both http: & https: urls https://github.com/Empact/roo/commit/fc5c5899d96dd5f9fbb68125d0efc8ce9be2c7e1

## [1.10.1] - 2011-11-14
### Fixed
- forgot dependency 'rubyzip'
- at least one external application does create xlsx-files with different internal file names which differ from the original file names of Excel. Solution: ignore lower-/upper case in file names.

## [1.10.0] - 2011-10-10
### Changed / Added
- New class Csv.
- Openoffice, Libreoffice: new method 'labels'
- Excelx: implemented all methods concerning labels
- Openoffice, Excelx: new methods concerning comments (comment, comment? and comments)

### Fixed
- XLSX: some cells were not recognized correctly from a spreadsheet file from a windows mobile phone.
- labels: Moved to a separate methode. There were problems if there was an access to a label before read_cells were called.

## [1.9.7] - 2011-08-27
### Fixed
- Openoffice: Better way for extracting formula strings, some characters were deleted at the formula string.

## [1.9.6] - 2011-08-03
### Changed / Added
- new class Libreoffice (Libreoffice should do exactly the same as the Openoffice
    class. It's just another name. Technically, Libreoffice is inherited from
    the Openoffice class with no new methods.

### Fixed
- Openoffice: file type check, deletion of temporary files not in ensure clause
- Cell type :datetime was not handled in the to_csv method
- Better deletion of temporary directories if something went wrong

## [1.9.5] - 2011-06-25
### Changed / Added
- Method #formulas moved to generic-spreadsheet class (the Excel version is
    overwritten because the spreadsheet gem currently does not support
    formulas.

### Fixed
- Openoffice/Excelx/Google: #formulas of an empty sheet should not result
    in an error message. Instead it should return an empty array.
- Openoffice/Excelx/Google: #to_yaml of an empty sheet should not result
    in an error message. Instead it should return an empty string.
- Openoffice/Excelx/Google: #to_matrix of an empty sheet should not result
    in an error message. Instead it should return an empty matrix.

## [1.9.4] - 2011-06-23
### Changed / Added
- removed gem 'builder'. Functionality goes to gem 'nokogiri'.

### Fixed
- Excel: remove temporary files if spreadsheed-file is not an excel file
    and an exception was raised
- Excelx: a referenced cell with a string had the content 0.0 not the
    correct string
- Fixed a problem with a date cell which was not recognized as a Date
    object (see 2011-05-21 in excelx.rb)

## [1.9.3] - 2010-02-12
### Changed / Added
- new method 'to_matrix'

### Fixed
- missing dependencies defined

## [1.9.2] - 2009-12-08
### Fixed
- double quoting of '"' fixed

## [1.9.1] - 2009-11-10
### Fixed
- syntax in nokogiri methods
- missing dependency ...rubyzip

## [1.9.0] - 2009-10-29
### Changed / Added
- Ruby 1.9 compatible
- oo.aa42 as a shortcut of oo.cell('aa',42)
- oo.aa42('sheet1') as a shortcut of oo.cell('aa',42,'sheet1')
- oo.anton as a reference to a cell labelled 'anton' (or any other label name)
    (currently only for Openoffice spreadsheets)

## [1.2.3] - 2009-01-04
### Fixed
- fixed encoding in #cell at exported Google-spreadsheets (.xls)

## [1.2.2] - 2008-12-14
### Changed / Added
- added celltype :datetime in Excelx
- added celltype :datetime in Google

## [1.2.1] - 2008-11-13
### Changed / Added
- added celltype :datetime in Openoffice and Excel

## [1.2.0] - 2008-08-24
### Changed / Added
- Excelx: improved the detection of cell type and conversion into roo types
- All: to_csv: changed boundaries from first_row,1..last_row,last_column to 1,1..last_row,last_column
- All: Environment variable "ROO_TMP" indicate where temporary directories will be created (if not set the default is the current working directory)

### Fixed
- Excel: improved the detection of last_row/last_column (parseexcel-gem bug?)
- Excel/Excelx/Openoffice: temporary directories were not removed at opening a file of the wrong type

## [1.1.0] - 2008-07-26
### Changed / Added
- Excel: speed improvements
- Changed the behavior of reading files with the wrong type

### Fixed
- Google: added normalize in set_value method
- Excel: last_row in Excel class did not work properly under some circumstances
- all: fixed a bug in #to_xml if there is an empty sheet

## [1.0.2] - 2008-07-04
### Fixed
- Excelx: fixed a bug when there are .xml.rels files in the XLSX archive
- Excelx: fixed a bug with celltype recognition (see comment with "2008-07-03")

## [1.0.1] - 2008-06-30
### Fixed
- Excel: row/column method Fixnum/Float confusion

## [1.0.0] - 2008-05-28
### Changed / Added
- support of Excel's new .xlsx file format
- method #to_xml for exporting a spreadsheet to an xml representation

### Fixed
- fixed a bug with excel-spreadsheet character conversion under Macintosh Darwin

## [0.9.4] - 2008-04-22
### Fixed
- fixed a bug with excel-spreadsheet character conversion under Solaris

## [0.9.3] - 2008-03-25
### Fixed
- no more tmp directories if an invalid spreadsheet file was openend

## [0.9.2] - 2008-03-24
### Changed / Added
- new celltype :time

### Fixed
- time values like '23:15' are handled as seconds from midnight

## [0.9.1] - 2008-03-23
### Changed / Added
- additional 'sheet' parameter in Google#set_value

### Fixed
- fixed a bug within Google#set_value. thanks to davecahill <dpcahill@gmail.com> for the patch.

## [0.9.0] - 2008-01-24
### Changed / Added
- better support of roo spreadsheets in rails views

## [0.8.5] - 2008-01-16
### Fixed
- fixed a bug within #to_cvs and explicit call of a sheet

## [0.8.4] - 2008-01-01
### Fixed
- fixed 'find_by_condition' for excel sheets (header_line= --> GenericSpredsheet)

## [0.8.3] - 2007-12-31
### Fixed
- another fix for the encoding issue in excel sheet-names
- reactived the Excel#find method which has been disappeared in the last restructoring, moved to GenericSpreadsheet

## [0.8.2] - 2007-12-28
### Changed / Added
- basename() only in method #info

### Fixed
- changed logging-method to mysql-database in test code with AR, table column 'class' => 'class_name'
- reactived the Excel#to_csv method which has been disappeared in the last restructoring

## [0.8.1] - 2007-12-22
### Fixed
- fixed a bug with first/last-row/column in empty sheet
- #info prints now '- empty -' if a sheet within a document is empty
- tried to fix the iconv conversion problem

## [0.8.0] - 2007-12-15
### Changed / Added
- Google online spreadsheets were implemented
- some methods common to more than one class were factored out to the GenericSpreadsheet (virtual) class

## [0.7.0] - 2007-11-23
### Changed / Added
- Openoffice/Excel: the most methods can be called with an option 'sheet'
    parameter which will be used instead of the default sheet
- Excel: improved the speed of CVS output
- Openoffice/Excel: new method #column
- Openoffice/Excel: new method #find
- Openoffice/Excel: new method #info
- better exception if a spreadsheet file does not exist

## [0.6.1] - 2007-10-06
### Changed / Added
- Openoffice: percentage-values are now treated as numbers (not strings)
- Openoffice: refactoring

### Fixed
- Openoffice: repeating date-values in a line are now handled correctly

## [0.6.0] - 2007-10-06
### Changed / Added
- csv-output to stdout or file

## [0.5.4] - 2007-08-27
### Fixed
- Openoffice: fixed a bug with internal representation of a spreadsheet (thanks to Ric Kamicar for the patch)

## [0.5.3] - 2007-08-26
### Changed / Added
- Openoffice: can now read zip-ed files
- Openoffice: can now read files from http://-URL over the net

## [0.5.2] - 2007-08-26
### Fixed
- excel: removed debugging output

## [0.5.1] - 2007-08-26
### Changed / Added
- Openoffice: Exception if an illegal sheet-name is selected
- Openoffice/Excel: no need to set a default_sheet if there is only one in
    the document
- Excel: can now read zip-ed files
- Excel: can now read files from http://-URL over the net

## [0.5.0] - 2007-07-20
### Changed / Added
- Excel-objects: the methods default_sheet= and sheets can now handle names instead of numbers
  ### Changedd the celltype methods to return symbols, not strings anymore (possible values are :formula, :float, :string, :date, :percentage (if you need strings you can convert it with .to_s)
- tests can now run on the client machine (not only my machine), if there are not public released files involved these tests are skipped

## [0.4.1] - 2007-06-27
### Fixed
- there was ONE false require-statement which led to misleading error messageswhen this gem was used

## [0.4.0] - 2007-06-27
### Changed / Added
- robustness: Exception if no default_sheet was set
- new method reload() implemented
- about 15 % more method documentation
- optimization: huge increase of speed (no need to use fixed borders anymore)
- added the method 'formulas' which gives you all formulas in a spreadsheet
- added the method 'set' which can set cells to a certain value
- added the method 'to_yaml' which can produce output for importing in a (rails) database

### Fixed
- ..row_as_letter methods were nonsense - removed
- @cells_read should be reset if the default_sheet is changed
- error in excel-part: strings are now converted to utf-8 (the parsexcel-gem gave me an error with my test data, which could not converted to .to_s using latin1 encoding)
- fixed bug when default_sheet is changed

## [0.3.0] - 2007-06-20
### Changed / Added
- Openoffice: formula support

## [0.2.7] - 2007-06-20
### Fixed
- Excel: float-numbers were truncated to integer

## [0.2.6] - 2007-06-19
### Fixed
- Openoffice: two or more consecutive cells with string content failed

## [0.2.5] - 2007-06-17
### Changed / Added
- Excel: row method implemented
- more tests

### Fixed
- Openoffice: row method fixed

## [0.2.4] - 2007-06-16
### Fixed
- ID 11605	Two cols with same value: crash roo (openoffice version only)

## [0.2.3] - 2007-06-02
### Changed / Added
- more robust call att Excel#default_sheet= when called with a name
- new method empty?
- refactoring

### Fixed
- bugfix in Excel#celltype
- bugfix (running under windows only) in closing the temp file before removing it

## [0.2.2] - 2007-06-01
### Fixed
- correct pathname for running with windows

## [0.2.2] - 2007-06-01
### Fixed
- incorrect dependencies fixed

## [0.2.0] - 2007-06-01
### Changed / Added
- support for MS-Excel Spreadsheets

## [0.1.2] - 2007-05-31
### Changed / Added
- cells with more than one character, like 'AA' can now be handled

## [0.1.1] - 2007-05-31
### Fixed
- bugfixes in first/last methods

## [0.1.0] - 2007-05-31
### Changed / Added
- new methods first/last row/column
- new method officeversion

## [0.0.3] - 2007-05-30
### Changed / Added
- new method row()

## [0.0.2] - 2007-05-30
### Changed / Added
- fixed some bugs
- more ways to access a cell

## [0.0.1] - 2007-05-25
### Changed / Added
- Initial release
