# patch for skipping blank rows in the case of
# having a spreadsheet with 30,000 nil rows appended
# to the actual data.  (it happens and your RAM will love me)
class Spreadsheet::Worksheet

  def each skip=dimensions[0]
    blanks = 0
    skip.upto(dimensions[1] - 1) do |i|
      if row(i).any?
        Proc.new.call(row(i))
      else
        blanks += 1
        blanks < 20 ? next : return
      end
    end
  end

end