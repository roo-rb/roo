# frozen_string_literal: true

require "spec_helper"

describe Roo::Excelx::SheetDoc do
  subject(:blank_children) { Roo::Excelx.new("test/files/blank_children.xlsx") }

  example "#last_row" do
    expect(subject.last_row).to eq 3
  end
end
