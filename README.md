# Roo

[![Build Status](https://img.shields.io/travis/roo-rb/roo.svg?style=flat-square)](https://travis-ci.org/roo-rb/roo) [![Code Climate](https://img.shields.io/codeclimate/github/roo-rb/roo.svg?style=flat-square)](https://codeclimate.com/github/roo-rb/roo) [![Coverage Status](https://img.shields.io/coveralls/roo-rb/roo.svg?style=flat-square)](https://coveralls.io/r/roo-rb/roo) [![Gem Version](https://img.shields.io/gem/v/roo.svg?style=flat-square)](https://rubygems.org/gems/roo)

Roo implements read access for all common spreadsheet types. It can handle:

* Excelx
* OpenOffice / LibreOffice
* CSV

## Additional Libraries

In addition, the [roo-xls](https://github.com/roo-rb/roo-xls) and [roo-google](https://github.com/roo-rb/roo-google) gems exist to extend Roo to support reading classic Excel formats (i.e. `.xls` and ``Excel2003XML``) and read/write access for Google spreadsheets.

# #Installation

Install as a gem

    $ gem install roo

Or add it to your Gemfile

```ruby
gem 'roo', '~> 2.0.0'
```
## Usage

Opening a spreadsheet

```ruby
require 'roo'

xlsx = Roo::Spreadsheet.open('./new_prices.xlsx')
xlsx = Roo::Excelx.new("./new_prices.xlsx")

# Use the extension option if the extension is ambiguous.
xlsx = Roo::Spreadsheet.open('./rails_temp_upload', extension: :xlsx)

xlsx.info
# => Returns basic info about the spreadsheet file
```

``Roo::Spreadsheet.open`` can accept both paths and ``File`` instances.

### Working with sheets

```ruby
ods.sheets
# => ['Info', 'Sheet 2', 'Sheet 3']   # an Array of sheet names in the workbook

ods.sheet('Info').row(1)
ods.sheet(0).row(1)

# Set the last sheet as the default sheet.
ods.default_sheet = ods.sheets.last
ods.default_sheet = s.sheets[3]
ods.default_sheet = 'Sheet 3'

# Iterate through each sheet
ods.each_with_pagename do |name, sheet|
  p sheet.row(1)
end
```

### Accessing rows and columns

Roo uses Excel's numbering for rows, columns and cells, so `1` is the first index, not `0` as it is in an ``Array``

```ruby
sheet.row(1)
# returns the first row of the spreadsheet.

sheet.column(1)
# returns the first column of the spreadsheet.
```

Almost all methods have an optional argument `sheet`. If this parameter is omitted, the default_sheet will be used.

```ruby
sheet.first_row(sheet.sheets[0])
# => 1             # the number of the first row
sheet.last_row
# => 42            # the number of the last row
sheet.first_column
# => 1             # the number of the first column
sheet.last_column
# => 10            # the number of the last column
```

#### Accessing cells

You can access the top-left cell in the following ways

```ruby
s.cell(1,1)
s.cell('A',1)
s.cell(1,'A')
s.a1

# Access the second sheet's top-left cell.
s.cell(1,'A',s.sheets[1])
```

#### Querying a spreadsheet
Use ``each`` with a ``block`` to iterate over each row.

If each is given a hash with the names of some columns, then each will generate a hash with the columns supplied for each row.

```ruby
sheet.each(id: 'ID', name: 'FULL_NAME') do |hash|
  puts hash.inspect
  # => { id: 1, name: 'John Smith' }
end
```

Use ``sheet.parse`` to return an array of rows. Column names can be a ``String`` or a ``Regexp``.

```ruby
sheet.parse(:id => /UPC|SKU/,:qty => /ATS*\sATP\s*QTY\z/)
# => [{:upc => 727880013358, :qty => 12}, ...]
```

Use the ``:header_search`` option to locate the header row and assign the header names.

```ruby
sheet.parse(header_search: [/UPC*SKU/,/ATS*\sATP\s*QTY\z/])
```

Use the ``:clean`` option to strip out control characters and surrounding white space.

```ruby
sheet.parse(:clean => true)
```

### Exporting spreadsheets
Roo has the ability to export sheets using the following formats. It
will only export the ``default_sheet``.

```ruby
sheet.to_csv
sheet.to_matrix
sheet.to_xml
sheet.to_yaml
```

### Excel (xlsx) Support

Stream rows from an Excelx spreadsheet.

```ruby
xlsx = Roo::Excelx.new("./test_data/test_small.xlsx")
xlsx.each_row_streaming do |row|
  puts row.inspect # Array of Excelx::Cell objects
end
```

Iterate over each row

```ruby
xlsx.each_row do |row|
  ...
end
```

``Roo::Excelx`` also provides these helpful methods.

```ruby
xlsx.excelx_type(3, 'C')
# => :numeric_or_formula

xlsx.cell(3, 'C')
# => 600000383.0

xlsx.excelx_value(row,col)
# => '0600000383'
```

``Roo::Excelx`` can access celltype, comments, font information, formulas, hyperlinks and labels.

```ruby
xlsx.comment(1,1, ods.sheets[-1])
xlsx.font(1,1).bold?
xlsx.formula('A', 2)
```

### OpenOffice / LibreOffice Support

Roo::OpenOffice supports for encrypted OpenOffice spreadsheets.

```ruby
# Load an encrypted OpenOffice Spreadsheet
ods = Roo::OpenOffice.new("myspreadsheet.ods", :password => "password")
```

``Roo::OpenOffice`` can access celltype, comments, font information, formulas and labels.

```ruby
ods.celltype
# => :percentage

ods.comment(1,1, ods.sheets[-1])

ods.font(1,1).italic?
# => false

ods.formula('A', 2)
```

### CSV Support

```ruby
# Load a CSV file
s = Roo::CSV.new("mycsv.csv")
```

Because Roo uses the [standard CSV library](), and you can use options available to that library to parse csv files. You can pass options using the ``csv_options`` key.

For instance, you can load tab-delimited files (``.tsv``), and you can use a particular encoding when opening the file.


```ruby
# Load a tab-delimited csv
s = Roo::CSV.new("mytsv.tsv", csv_options: {col_sep: "\t"})

# Load a csv with an explicit encoding
s = Roo::CSV.new("mycsv.csv", csv_options: {encoding: Encoding::ISO_8859_1})
```

## Upgrading from Roo 1.13.x
If you use ``.xls`` or Google spreadsheets, you will need to install ``roo-xls`` or ``roo-google`` to continue using that functionality.

Roo's public methods have stayed relatively consistent between 1.13.x and 2.0.0, but please check the [Changelog](https://github.com/roo-rb/roo/blob/master/CHANGELOG.md) to better understand the changes made since 1.13.x.



## Contributing
### Features
1. Fork it ( https://github.com/[my-github-username]/roo/fork )
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'My new feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create a new Pull Request

### Issues

If you find an issue, please create a gist and refer to it in an issue ([sample gist](https://gist.github.com/stevendaniels/98a05849036e99bb8b3c)). Here are some instructions for creating such a gist.

1. [Create a gist](https://gist.github.com) with code that creates the error.
2. Clone the gist repo locally, add a stripped down version of the offending spreadsheet to the gist repo, and push the gist's changes master.
3. Paste the gist url here.


## License
[Roo uses an MIT License](https://github.com/roo-rb/roo/blob/master/LICENSE)
