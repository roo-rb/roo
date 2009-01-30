require 'rubygems'
require 'roo'

#-- create a new spreadsheet within your google-spreadsheets and paste
#-- the 'key' parameter in the spreadsheet URL 
MAXTRIES = 1000
print "what's your name? "
my_name = gets.chomp
print "where do you live? "
my_location = gets.chomp
print "your message? (if left blank, only your name and location will be inserted) "
my_message = gets.chomp
spreadsheet = Google.new('ptu6bbahNZpY0N0RrxQbWdw')
spreadsheet.default_sheet = 'Sheet1'
success = false
MAXTRIES.times do
  col = rand(10)+1
  row = rand(10)+1
  if spreadsheet.empty?(row,col)
    if my_message.empty?
      text = Time.now.to_s+" "+"Greetings from #{my_name} (#{my_location})"
    else
      text = Time.now.to_s+" "+"#{my_message} from #{my_name} (#{my_location})"
    end
    spreadsheet.set_value(row,col,text)
    puts "message written to row #{row}, column #{col}"
    success = true
    break
  end
  puts "Row #{row}, column #{col} already occupied, trying again..."
end
puts "no empty cell found within #{MAXTRIES} tries" if !success

