require_relative '../spec_helper'

describe Table do
  
  before(:all) do
    @table = Table.new("FirstTable",:sheet1,"B3:D7","1","ColA","ColB","ColC")
  end
    
  it 'should be able to return a reference for [:table_reference, "FirstTable", "ColA"]]' do
    @table.reference_for("FirstTable", "ColA",:sheet1,"A1").should == [:sheet_reference, :sheet1, [:area, :B4, :B6]]
  end

  it 'should return a :ref error if table column does not exist' do
    @table.reference_for("FirstTable", "NotAColumn",:sheet1,"A1").should == [:error,"#REF!"]
    @table.reference_for("FirstTable", "[ColA]:[NotAColumn]",:sheet1,"A1").should == [:error,"#REF!"]
    @table.reference_for("FirstTable", "[#Headers],[NotAColumn]",:sheet1,"F5").should == [:error,"#REF!"]
    @table.reference_for("FirstTable", "[#Totals],[NotAColumn]",:sheet1,"F5").should == [:error,"#REF!"]
    @table.reference_for("FirstTable", "[#This Row],[NotAColumn]",:sheet1,"F5").should == [:error,"#REF!"]
  end

  it 'should be able to return a reference for [:table_reference, "FirstTable", "ColA"]] from a cell within the table' do
    @table.reference_for("FirstTable", "ColA",:sheet1,:C5).should == [:sheet_reference, :sheet1, [:cell, :B5]]
  end

  it 'should be able to return a reference for [:table_reference, "FirstTable", "ColA"]] from a cell within the total row' do
    @table.reference_for("FirstTable", "ColA",:sheet1,:B7).should == [:sheet_reference, :sheet1, [:area, :B4, :B6]]
  end  

  it 'should be able to return a reference for [:table_reference, "FirstTable", "#Totals"]]' do
    @table.reference_for("FirstTable", "#Totals",:sheet1,"A1").should == [:sheet_reference, :sheet1, [:area, :B7, :D7]]
  end

  it 'should be able to return a reference for [:table_reference, "FirstTable", "#Totals"]] from a cell within the table' do
    @table.reference_for("FirstTable", "#Totals",:sheet1,:C5).should == [:sheet_reference, :sheet1, [:cell, :C7]]
  end

  it 'should be able to return a reference for [:table_reference, "FirstTable", "#Headers"]]' do
    @table.reference_for("FirstTable", "#Headers",:sheet1,"A1").should == [:sheet_reference, :sheet1, [:area, :B3, :D3]]
  end

  it 'should be able to return a reference for [:table_reference, "FirstTable", "#Headers"]] from a cell within the table' do
    @table.reference_for("FirstTable", "#Headers",:sheet1,:C5).should == [:sheet_reference, :sheet1, [:cell, :C3]]
  end

  it 'should be able to return a reference for [:table_reference, "FirstTable", "#Data"]]' do
    @table.reference_for("FirstTable", "#Data",:sheet1,"A1").should == [:sheet_reference, :sheet1, [:area, :B4, :D6]]
  end

  it 'should be able to return a reference for [:table_reference, "FirstTable", ""]] which should be the same as a reference to #Data' do
    @table.reference_for("FirstTable","",:sheet1,"A1").should == [:sheet_reference, :sheet1, [:area, :B4, :D6]]
    # @table.reference_for("FirstTable", "",:sheet1,"A1").should == [:sheet_reference, :sheet1, [:area, :B3, :D7]]
  end
  
  it 'should be able to return a reference for [:table_reference, "FirstTable", "#All"]]' do
    @table.reference_for("FirstTable", "#All",:sheet1,"A1").should == [:sheet_reference, :sheet1, [:area, :B3, :D7]]
    
  end

  it 'should be able to return a reference for [:table_reference, "FirstTable", "#This Row"]]' do
    @table.reference_for("FirstTable", "#This Row",:sheet1,"F5").should == [:sheet_reference, :sheet1, [:area, :B5, :D5]]
  end
  
  it 'should be able to return a reference for [:table_reference, "FirstTable", "[#Headers],[ColB]"]' do
    @table.reference_for("FirstTable", "[#Headers],[ColB]",:sheet1,"F5").should == [:sheet_reference, :sheet1, [:cell, :C3]]
  end

  it 'should be able to return a reference for [:table_reference, "FirstTable", "[#Totals],[ColA]"]' do
    @table.reference_for("FirstTable", "[#Totals],[ColA]",:sheet1,"F5").should == [:sheet_reference, :sheet1, [:cell, :B7]]
  end
  
  it 'should be able to return a reference for [:table_reference, "FirstTable", "[#This Row],[ColA]"]' do
    @table.reference_for("FirstTable", "[#This Row],[ColA]",:sheet1,"F5").should == [:sheet_reference, :sheet1, [:cell, :B5]]
  end
  
  it 'should be able to return a reference for [:table_reference, "FirstTable", "[#Headers],[ColB]:[ColC]"]' do
    @table.reference_for("FirstTable", "[#Headers],[ColB]:[ColC]",:sheet1,"F5").should == [:sheet_reference, :sheet1, [:area, :C3,:D3]]
  end

  it 'should be able to return a reference for [:table_reference, "FirstTable", "[#Totals],[ColA]:[ColC]"]' do
    @table.reference_for("FirstTable", "[#Totals],[ColA]:[ColC]",:sheet1,"F5").should == [:sheet_reference, :sheet1, [:area, :B7,:D7]]
  end

  it 'should be able to return a reference for [:table_reference, "FirstTable", "[#This Row],[ColB]:[ColC]"]' do
    @table.reference_for("FirstTable", "[#This Row],[ColB]:[ColC]",:sheet1,"F5").should == [:sheet_reference, :sheet1, [:area, :C5,:D5]]
  end 
  
  it "should be able to tell whether a reference is inside or outside of the table" do 
    @table.includes?(:sheet1,:B3).should == true
    @table.includes?(:sheet1,:D7).should == true
    @table.includes?(:sheet1,:B4).should == true
    @table.includes?(:sheet1,"C6").should == true
    
    @table.includes?("sheet2",:B3).should == false
    @table.includes?(:sheet1,"A3").should == false
    @table.includes?(:sheet1,"D8").should == false
  end
end

describe "Table bug, not returning reference for whole table" do
  
 before(:all) do
    @table = Table.new("Global.Assumptions.Energy.Prices.High","Global assumptions","D84:N90","0","Fuel","Unit","2010", "2015", "2020", "2025", "2030", "2035", "2040", "2045", "2050")
  end
    

 it "Should return the area of the data on the table when passed an empty structured reference" do
    @table.reference_for("Global.Assumptions.Energy.Prices.High", "", "Sheet1", "A1").should == [:sheet_reference, :"Global assumptions", [:area, :D85, :N90]]

 end
 

end
