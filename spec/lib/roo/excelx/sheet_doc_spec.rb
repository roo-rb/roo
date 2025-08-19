# frozen_string_literal: true

require "spec_helper"

describe Roo::Excelx::SheetDoc do
  subject(:blank_children) { Roo::Excelx.new("test/files/blank_children.xlsx") }

  example "#last_row" do
    expect(subject.last_row).to eq 3
  end

  describe "#each_cell" do
    let(:relationships) { instance_double(Roo::Excelx::Relationships, include_type?: false) }
    let(:styles) { instance_double(Roo::Excelx::Styles, style_format: 'General') }
    let(:shared) { instance_double(Roo::Excelx::Shared, styles: styles) }
    let(:sheet_doc) { described_class.new(nil, relationships, shared) }

    context "empty v element" do
      let(:row_xml) { Nokogiri.parse('<row r="1"><c r="A1"><v/></c></row>').root }

      it "returns an empty cell" do
        expect(sheet_doc.each_cell(row_xml)).to all(be_a(Roo::Excelx::Cell::Empty))
      end
    end

    context "no v element" do
      let(:row_xml) { Nokogiri.parse('<row r="1"><c r="A1"></c></row>').root }

      it "returns an empty cell" do
        expect(sheet_doc.each_cell(row_xml)).to all(be_a(Roo::Excelx::Cell::Empty))
      end
    end
  end
end
