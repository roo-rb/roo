require 'roo/excelx/extractor'

module Roo
  class Excelx::Comments < Excelx::Extractor
    def comments
      @comments ||=
        if doc_exists?
          Hash[doc.xpath("//comments/commentList/comment").map do |comment|
            [Roo::Base.ref_to_key(comment.attributes['ref'].to_s), comment.at_xpath('./text/r/t').text]
          end]
        else
          {}
        end
    end
  end
end
