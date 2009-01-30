require 'rubygems'
require 'roo'
require 'soap/rpc/standaloneServer'

NS = "spreadsheetserver" # name of your service = namespace
class Server2 < SOAP::RPC::StandaloneServer

  def on_init
    spreadsheet = Openoffice.new("./Ferien-de.ods")
    add_method(spreadsheet, 'cell', 'row', 'col')
    add_method(spreadsheet, 'officeversion')
    add_method(spreadsheet, 'first_row')
    add_method(spreadsheet, 'last_row')
    add_method(spreadsheet, 'first_column')
    add_method(spreadsheet, 'last_column')
    add_method(spreadsheet, 'sheets')
    #add_method(spreadsheet, 'default_sheet=', 's')
    # method with '...=' did not work? alias method 'set_default_sheet' created
    add_method(spreadsheet, 'set_default_sheet', 's')
  end

end

PORT = 12321
puts "serving at port #{PORT}"
svr = Server2.new('Roo', NS, '0.0.0.0', PORT)

trap('INT') { svr.shutdown }
svr.start
