class Roo::Spreadsheet::HTML
  
  def initialize(filename)
    raw = File.open(filename).read
    table = raw.match(/<table>(.*)<\/table>/)[1] rescue raise('No <table> found.')
    @doc = Nokogiri::HTML.open(table)
    @cell = {}
    parse_html
  end
  
  def parse_html
    @doc.xpath('//tr').each_with_index do |row,row_no|
      row.children.each_with_index do |col,col_no|
        @cell[[row_no,col_no]] = col.text
      end
    end
  end
  
end