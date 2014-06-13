require_relative '../spec_helper'

describe ExcelToX do
  
  it "Should allow named_references_to_keep to be specified as a block" do
    c = ExcelToX.new
    c.named_references_to_keep = lambda { |named_reference| named_reference =~ /^keep/ }
    c.instance_variable_set("@named_references", { :keeper => :blah, :looser => :blah })
    c.instance_variable_set("@table_areas", { })
    c.clean_named_references_to_keep
    c.named_references_to_keep.should == [ :keeper ]
  end

  it "Should allow named_references_that_can_be_set_at_runtime to be specified as a block" do
    c = ExcelToX.new
    c.named_references_that_can_be_set_at_runtime = lambda { |named_reference| named_reference =~ /^setter/ }
    c.instance_variable_set("@named_references", { :setterone => :blah, :looser => :blah })
    c.instance_variable_set("@table_areas", { })
    c.clean_named_references_that_can_be_set_at_runtime
    c.named_references_that_can_be_set_at_runtime.should == [ :setterone ]
  end

  it "Should allow named_references_to_keep to inlcude table names" do
    c = ExcelToX.new
    c.named_references_to_keep = lambda { |named_reference| named_reference =~ /^keep/ }
    c.instance_variable_set("@named_references", { :keeper => :blah, :looser => :blah })
    c.instance_variable_set("@table_areas", {:keep_table => :ref, :loose_table => :ref})
    c.clean_named_references_to_keep
    c.named_references_to_keep.should == [ :keeper, :keep_table ]
  end
end
