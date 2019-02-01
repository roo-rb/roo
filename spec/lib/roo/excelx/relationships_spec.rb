# frozen_string_literal: true

require "spec_helper"

describe Roo::Excelx::Relationships do
  subject(:relationships) { Roo::Excelx::Relationships.new Roo::Excelx.new(path).rels_files[0] }

  describe "#include_type?" do
    [
      ["with hyperlink type", "test/files/link.xlsx", true, false],
      ["with nil path", "test/files/Bibelbund.xlsx", false, false],
      ["with comments type", "test/files/comments-google.xlsx", false, true],
    ].each do |context_desc, file_path, hyperlink_value, comments_value|
      context context_desc do
        let(:path) { file_path }

        it "should return #{hyperlink_value} for hyperlink" do
          expect(subject.include_type?("hyperlink")).to be hyperlink_value
        end

        it "should return #{hyperlink_value} for link" do
          expect(subject.include_type?("link")).to be hyperlink_value
        end

        it "should return false for hypelink" do
          expect(subject.include_type?("hypelink")).to be false
        end

        it "should return false for coment" do
          expect(subject.include_type?("coment")).to be false
        end

        it "should return #{comments_value} for comments" do
          expect(subject.include_type?("comments")).to be comments_value
        end

        it "should return #{comments_value} for comment" do
          expect(subject.include_type?("comment")).to be comments_value
        end
      end
    end
  end
end
