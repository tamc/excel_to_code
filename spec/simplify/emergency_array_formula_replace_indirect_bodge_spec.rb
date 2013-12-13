require_relative '../spec_helper'

describe EmergencyArrayFormulaReplaceIndirectBodge do

  it "should replace INDIRECT() functions in array formulae with the reference that they refer to, even if they haven't been passed a string value. THIS IS A BODGE" do

    r = EmergencyArrayFormulaReplaceIndirectBodge.new
    r.current_sheet_name = :Sheet1
    r.references = {
      [:Sheet1, :A1] => [:string, "Sheet1!"]
    }
    r.replace([:function, :INDIRECT, [:string, "$A$5"]]).should == [:cell, :'$A$5']
    r.replace([:function, :INDIRECT, [:string_join, [:string, "Sheet1!"], [:string, "$A$5"]]]).should == [:sheet_reference, :Sheet1, [:cell, :'$A$5']]
    r.replace([:function, :INDIRECT, [:string_join, [:cell, :A1], [:string, "$A$5"]]]).should == [:sheet_reference, :Sheet1, [:cell, :'$A$5']]

  end # / it


end # / describe
