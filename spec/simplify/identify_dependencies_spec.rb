require_relative '../spec_helper'

describe IdentifyDependencies do
  
it "should recursively work through formulae, identifying dependencies" do
  references = {
    [:sheet1, :A1] => [:cell, :"$A$2"],
    [:sheet1, :A2] => [:cell, :A3],
    [:sheet1, :A3] => [:number, 1],
    [:sheet2, :A1] => [:cell, :A2],
    [:sheet2, :A2] => [:sheet_reference,:sheet3,[:cell,:A1]],
    [:sheet3, :A1] => [:number, 5],
    [:sheet3, :A5] => [:number, 10]    
  }

dependencies = {
  :sheet1 => {
    :A1 => true,
    :A2 => true,
    :A3 => true
  }
}

identifier = IdentifyDependencies.new
identifier.references = references
identifier.add_depedencies_for(:sheet1,:A1)
identifier.dependencies.should == dependencies
end # / do

it "should be able to add all dependencies for a sheet" do
references = {
    [:sheet1, :A1] => [:cell, :"$A$2"],
    [:sheet1, :A2] => [:cell, :A3],
    [:sheet1, :A3] => [:number, 1],
    [:sheet2, :A1] => [:cell, :A2],
    [:sheet2, :A2] => [:sheet_reference,:sheet3,[:cell,:A1]],
    [:sheet3, :A1] => [:number, 5],
    [:sheet3, :A5] => [:number, 10]    
}

dependencies = {
  :sheet2 => {
    :A1 => true,
    :A2 => true,
  },
  :sheet3 => {
    :A1 => true
  }
}

identifier = IdentifyDependencies.new
identifier.references = references
identifier.add_depedencies_for(:sheet2)
identifier.dependencies.should == dependencies
end # / do


end # / describe
