require 'roo/excelx/extractor'

module Roo
  class Excelx::SharedStrings < Excelx::Extractor
    def [](index)
      to_a[index]
    end

    def to_a
      @array ||= extract_shared_strings
    end

    private

    def extract_shared_strings
      if doc_exists?
        # read the shared strings xml document
        doc.xpath("/sst/si").map do |si|
          shared_string = ''
          si.children.each do |elem|
            case elem.name
              when 'r'
                elem.children.each do |r_elem|
                  if r_elem.name == 't'
                    shared_string << r_elem.content
                  end
                end
              when 't'
                shared_string = elem.content
            end
          end
          shared_string
        end
      else
        []
      end
    end

  end
end
