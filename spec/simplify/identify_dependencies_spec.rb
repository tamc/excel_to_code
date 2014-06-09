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

  it "should raise an error if a circular reference is detected" do
    references = {
      [:sheet2, :A1] => [:cell, :A2],
      [:sheet2, :A2] => [:sheet_reference,:sheet3,[:cell,:A1]],
      [:sheet3, :A1] => [:sheet_reference, :sheet2, [:cell, :A2]],
    }
    identifier = IdentifyDependencies.new
    identifier.references = references
    expect { identifier.add_depedencies_for(:sheet2) }.to raise_error(ExcelToCodeException)
  end

  it "should not be dumb in its circular reference checks" do
    references = {
      [:sheet2, :A1] => [:cell, :A2],
      [:sheet2, :A2] => [:arithmetic, [:sheet_reference,:sheet3,[:cell,:A1]], [:operator, :+], [:sheet_reference, :sheet3, [:cell, :A1]]],
      [:sheet3, :A1] => [:number, 10],
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
  end

  it "should not be dumb in its circular reference checks" do
    references = {
      [:'XII.b', :F228] => [:arithmetic, [:function, :INDEX, [:array, [:row, [:number, 0.0]], [:row, [:number, 0.0]], [:row, [:number, 0.0]], [:row, [:number, 0.0]]], [:function, :MATCH, [:sheet_reference, :Control, [:cell, :E33]], [:array, [:row, [:number, 1.0]], [:row, [:number, 2.0]], [:row, [:number, 3.0]], [:row, [:number, 4.0]]], [:number, 0.0]]], [:operator, :*], [:function, :INDEX, [:array, [:row, [:inlined_blank]], [:row, [:inlined_blank]], [:row, [:inlined_blank]], [:row, [:inlined_blank]]], [:function, :MATCH, [:sheet_reference, :Control, [:cell, :E33]], [:array, [:row, [:number, 1.0]], [:row, [:number, 2.0]], [:row, [:number, 3.0]], [:row, [:number, 4.0]]], [:number, 0.0]]]],
      [:'XII.b', :F348] => [:arithmetic, [:sheet_reference, :"XII.b", [:cell, :F223]], [:operator, :+], [:sheet_reference, :"XII.b", [:cell, :F228]]],
      [:'2007', :J25] => [:function, :IFERROR, [:sheet_reference, :"XII.b", [:cell, :F348]], [:number, 0.0]],
      [:'2007', :J32] => [:function, :ENSURE_IS_NUMBER, [:arithmetic, [:sheet_reference, :"2007", [:cell, :J24]], [:operator, :+], [:sheet_reference, :"2007", [:cell, :J25]]]],
      [:'2007', :J109] => [:function, :ENSURE_IS_NUMBER, [:function, :ENSURE_IS_NUMBER, [:sheet_reference, :"2007", [:cell, :J32]]]],
      [:'Intermediate output', :AY7] => [:sheet_reference, :"2007", [:cell, :J109]]
    }
    dependencies = {
      :"2007" => {:J109=>true, :J32=>true, :J24=>true, :J25=>true},
      :Control => {:E33=>true},
      :"Intermediate output" => {:AY7=>true},
      :"XII.b" => {:F348=>true, :F223=>true, :F228=>true}
    }

    identifier = IdentifyDependencies.new
    identifier.references = references
    identifier.add_depedencies_for(:'Intermediate output')
    identifier.dependencies.should == dependencies
  end


end # / describe
