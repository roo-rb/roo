require "test_helper"

class TestAttrReaderDefault < Minitest::Test
  def base
    Roo::Excelx::Cell::Base
  end

  def boolean
    Roo::Excelx::Cell::Boolean
  end

  def class_date
    Roo::Excelx::Cell::Date
  end

  def datetime
    Roo::Excelx::Cell::DateTime
  end

  def empty
    Roo::Excelx::Cell::Empty
  end

  def number
    Roo::Excelx::Cell::Number
  end

  def string
    Roo::Excelx::Cell::String
  end

  def base_date
    ::Date.new(1899, 12, 30)
  end

  def base_timestamp
    ::Date.new(1899, 12, 30).to_datetime.to_time.to_i
  end

  def class_time
    Roo::Excelx::Cell::Time
  end

  def test_cell_default_values
    assert_values base.new(nil, nil, [], 1, nil, nil), default_type: :base, :@default_type => nil, style: 1, :@style => nil
    assert_values boolean.new("1", nil, nil, nil, nil), default_type: :boolean, :@default_type => nil, cell_type: :boolean, :@cell_type => nil
    assert_values class_date.new("41791", nil, [:numeric_or_formula, "mm-dd-yy"], 6, nil, base_date, nil), default_type: :date, :@default_type => nil
    assert_values class_time.new("0.521", nil, [:numeric_or_formula, "hh:mm"], 6, nil, base_timestamp, nil), default_type: :time, :@default_type => nil
    assert_values datetime.new("41791.521", nil, [:numeric_or_formula, "mm-dd-yy hh:mm"], 6, nil, base_timestamp, nil), default_type: :datetime, :@default_type => nil
    assert_values empty.new(nil), default_type: nil, :@default_type => nil, style: nil, :@style => nil
    assert_values number.new("42", nil, ["0"], nil, nil, nil), default_type: :float, :@default_type => nil
    assert_values string.new("1", nil, nil, nil, nil), default_type: :string, :@default_type => nil, cell_type: :string, :@cell_type => nil

    assert_values base.new(nil, nil, [], 2, nil, nil), style: 2, :@style => 2
  end

  def assert_values(object, value_hash)
    value_hash.each do |attr_name, expected_value|
      value = if attr_name.to_s.include?("@")
                object.instance_variable_defined?(attr_name) ? object.instance_variable_get(attr_name) : nil
              else
                object.public_send(attr_name)
      end

      if expected_value
        assert_equal expected_value, value
      else
        assert_nil value
      end
    end
  end
end
