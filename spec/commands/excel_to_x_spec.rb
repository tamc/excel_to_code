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

  it "Should have a c_name_for method that converts names into C compatible variants" do
    c = ExcelToX.new
    c.c_name_for("sheet1").should == "sheet1"
    c.c_name_for("A name with (unaceptable) characters").should == "a_name_with_unaceptable_characters"
    c.c_name_for("A clashing name").should == "a_clashing_name"
    c.c_name_for("A (clashing) name").should == "a_clashing_name2"
    c.c_name_for("A [clashing] name").should == "a_clashing_name3"
    c.c_name_for("a [clashing]  name").should == "a_clashing_name4"
    c.c_name_for("2010").should == "s2010"
    c.c_name_for("(Not appropriate!)").should == "s_not_appropriate_"
  end
end
