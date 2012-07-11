require File.dirname(__FILE__) + '/test_helper.rb'

class TestGenericSpreadsheet < Test::Unit::TestCase
  def test_letters
    assert_equal 1, Roo::GenericSpreadsheet.letter_to_number('A')
    assert_equal 1, Roo::GenericSpreadsheet.letter_to_number('a')
    assert_equal 2, Roo::GenericSpreadsheet.letter_to_number('B')
    assert_equal 26, Roo::GenericSpreadsheet.letter_to_number('Z')
    assert_equal 27, Roo::GenericSpreadsheet.letter_to_number('AA')
    assert_equal 27, Roo::GenericSpreadsheet.letter_to_number('aA')
    assert_equal 27, Roo::GenericSpreadsheet.letter_to_number('Aa')
    assert_equal 27, Roo::GenericSpreadsheet.letter_to_number('aa')
  end

  
end