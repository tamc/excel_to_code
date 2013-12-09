require_relative '../spec_helper'

describe ReplaceSharedStrings do
  
  it "should take the results of extract_worksheet_names.rb and extract_relationships.rb and return one line per worksheet: worksheet name then tab then worksheet filename" do
    shared_strings = [[:string, "One"], [:string, "Two"], [:string, "Three"]]
    values = StringIO.new("A2\t[:shared_string, \"0\"]\n")
    output = StringIO.new
    ReplaceSharedStrings.replace(values,shared_strings,output)
    output.string.should == "A2\t[:string, \"One\"]\n"
  end
end
