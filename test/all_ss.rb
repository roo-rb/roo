require 'roo'
Dir.glob("test/**/*").each do |fn|
	if fn.end_with? '.ods'
		begin
			oo = Openoffice.new fn
			print File.basename(fn) + " "
			puts oo.officeversion
		rescue #Zip::ZipError
			# file is not a real .ods spreadsheet file
			#puts "not an Openoffice-spreadsheet"
		end
	end
end
