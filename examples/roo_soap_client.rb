require 'soap/rpc/driver'

  def ferien_fuer_region(proxy, region, year=nil)
    proxy.first_row.upto(proxy.last_row) { |row|
      if proxy.cell(row, 2) == region
        jahr = proxy.cell(row,1).to_i
        if year == nil || jahr == year
          bis_datum = proxy.cell(row,5)
          if DateTime.now > bis_datum
            print '('
          end
          print jahr.to_s+" "
          print proxy.cell(row,2)+" "
          print proxy.cell(row,3)+" "
          print proxy.cell(row,4).to_s+" "
          print bis_datum.to_s+" "
          print (proxy.cell(row,6) || '')+" "
          if DateTime.now > bis_datum
            print ')'
          end
          puts
        end
      end 
    } 
  end

proxy = SOAP::RPC::Driver.new("http://localhost:12321","spreadsheetserver")
proxy.add_method('cell','row','col')
proxy.add_method('officeversion')
proxy.add_method('last_row')
proxy.add_method('last_column')
proxy.add_method('first_row')
proxy.add_method('first_column')
proxy.add_method('sheets')
proxy.add_method('set_default_sheet','s')
proxy.add_method('ferien_fuer_region', 'region')

sheets = proxy.sheets
proxy.set_default_sheet(sheets.first)

puts "first row: #{proxy.first_row}"
puts "first column: #{proxy.first_column}"
puts "last row: #{proxy.last_row}"
puts "last column: #{proxy.last_column}"
puts "cell: #{proxy.cell('C',8)}"
puts "cell: #{proxy.cell('F',12)}"
puts "officeversion: #{proxy.officeversion}"
puts "Berlin:"

ferien_fuer_region(proxy, "Berlin")



