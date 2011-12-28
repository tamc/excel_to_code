require_relative '../spec_helper'
require_relative '../../src/rewrite/rewrite_values_to_include_shared_strings'
require 'stringio'

describe RewriteValuesToIncludeSharedStrings do
  
  it "should take the results of extract_worksheet_names.rb and extract_relationships.rb and return one line per worksheet: worksheet name then tab then worksheet filename" do
    shared_strings = StringIO.new("One\nTwo\nThree\n")
    values = StringIO.new("A1\tn\t10\nA2\ts\t0\nA3\tstr\tNot Shared\n")
    output = StringIO.new
    RewriteValuesToIncludeSharedStrings.rewrite(values,shared_strings,output)
    output.string.should == "A1\tn\t10\nA2\tstr\tOne\nA3\tstr\tNot Shared\n"
  end
end
