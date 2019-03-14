require "test_helper"

class TestImageIndex < Minitest::Test
  def test_number_of_images
    with_each_spreadsheet(name: "kangaroos", format: [:excelx]) do |oo|
      assert_equal 4, oo.sheet_for(0).images.size
      assert_equal 0, oo.sheet_for(1).images.size
      assert_equal 1, oo.sheet_for(2).images.size
    end
  end

  def test_order_of_images
    with_each_spreadsheet(name: "kangaroos", format: [:excelx]) do |oo|
      expected = {"rId1"=>"roo_media_image1.jpeg", "rId2"=>"roo_media_image2.jpeg", "rId3"=>"roo_media_image3.jpeg", "rId4"=>"roo_media_image4.jpeg"}

      assert_equal expected.keys, oo.sheet_for(0).images.keys
      assert_equal expected.values, oo.sheet_for(0).images.values
    end
  end
end
