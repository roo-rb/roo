def spreadsheet(spreadsheet, sheets, options={})
  @rspreadsheet = spreadsheet
  coordinates = true # default
  o=""
  if options[:coordinates] != nil
    o << ":coordinates uebergeben: #{options[:coordinates]}"
    coordinates = options[:coordinates]
  end
  if options[:bgcolor]
    bgcolor = options[:bgcolor]
  else
    bgcolor = false
  end

  sheets.each { |sheet|
    @rspreadsheet.default_sheet = sheet
    linenumber = @rspreadsheet.first_row(sheet) 
    if options[:first_row]
      linenumber += (options[:first_row]-1) 
    end
    o << '<table border="0" cellspacing="1" cellpadding="1">'
    if options[:first_row]
      first_row = options[:first_row]
    end
    if options[:last_row]
      last_row = options[:last_row]
    end
    if options[:first_column]
      first_column = options[:first_column]
    end
    if options[:last_column]
      last_column = options[:last_column]
    end
    first_row    = @rspreadsheet.first_row(sheet) unless first_row
    last_row     = @rspreadsheet.last_row(sheet) unless last_row
    first_column = @rspreadsheet.first_column(sheet) unless first_column
    last_column  = @rspreadsheet.last_column(sheet) unless last_column
    if coordinates
      o << "  <tr>"
      o << "  <td>&nbsp;</td>"
      @rspreadsheet.first_column(sheet).upto(@rspreadsheet.last_column(sheet)) {|c| 
        if c < first_column or c > last_column
          next
        end
        o << "    <td>"
        o << "      <b>#{GenericSpreadsheet.number_to_letter(c)}</b>"
        o << "    </td>"
      } 
      o << "</tr>"
    end
    @rspreadsheet.first_row.upto(@rspreadsheet.last_row) do |y|
      if first_row and (y < first_row or y > last_row)
        next
      end
      o << "<tr>"
      if coordinates
        o << "<td><b>#{linenumber.to_s}</b></td>"
      end
      linenumber += 1
      @rspreadsheet.first_column(sheet).upto(@rspreadsheet.last_column(sheet)) do |x|
        if x < first_column or x > last_column
          next
        end
        if bgcolor
          o << "<td bgcolor=\"#{bgcolor}\">"
        else
          o << '<td bgcolor="lightgreen">'
        end
        if @rspreadsheet.cell(y,x).to_s.empty?
          o << "&nbsp;"
        else
          o << "#{@rspreadsheet.cell(y,x)}"
        end
        o << "</td>"
      end
      o << "</tr>"
    end
    o << "</table>"
  } # each sheet
  return o
end

