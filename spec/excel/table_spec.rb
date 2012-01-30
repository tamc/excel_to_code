require_relative '../spec_helper'

describe Table do
  
  before(:all) do
    @table = Table.new("FirstTable","sheet1","B3:D7","1","ColA","ColB","ColC")
  end
    
  it 'should be able to return a reference for [:table_reference, "FirstTable", "ColA"]]' do
    @table.reference_for("FirstTable", "ColA","sheet1","A1").should == [:sheet_reference, "sheet1", [:area, "B4", "B6"]]
  end

  it 'should be able to return a reference for [:table_reference, "FirstTable", "ColA"]] from a cell within the table' do
    @table.reference_for("FirstTable", "ColA","sheet1","C5").should == [:sheet_reference, "sheet1", [:cell, "B5"]]
  end

  it 'should be able to return a reference for [:table_reference, "FirstTable", "#Totals"]]' do
    @table.reference_for("FirstTable", "#Totals","sheet1","A1").should == [:sheet_reference, "sheet1", [:area, "B7", "D7"]]
  end

  it 'should be able to return a reference for [:table_reference, "FirstTable", "#Headers"]]' do
    @table.reference_for("FirstTable", "#Headers","sheet1","A1").should == [:sheet_reference, "sheet1", [:area, "B3", "D3"]]
  end

  it 'should be able to return a reference for [:table_reference, "FirstTable", "#Data"]]' do
    @table.reference_for("FirstTable", "#Data","sheet1","A1").should == [:sheet_reference, "sheet1", [:area, "B4", "D6"]]
  end

  it 'should be able to return a reference for [:table_reference, "FirstTable", ""]] which should be the same as a reference to #Data' do
    @table.reference_for("FirstTable","","sheet1","A1").should == [:sheet_reference, "sheet1", [:area, "B4", "D6"]]
    # @table.reference_for("FirstTable", "","sheet1","A1").should == [:sheet_reference, "sheet1", [:area, "B3", "D7"]]
  end
  
  it 'should be able to return a reference for [:table_reference, "FirstTable", "#All"]]' do
    @table.reference_for("FirstTable", "#All","sheet1","A1").should == [:sheet_reference, "sheet1", [:area, "B3", "D7"]]
    
  end

  it 'should be able to return a reference for [:table_reference, "FirstTable", "#This Row"]]' do
    @table.reference_for("FirstTable", "#This Row","sheet1","F5").should == [:sheet_reference, "sheet1", [:area, "B5", "D5"]]
  end
  
  it 'should be able to return a reference for [:table_reference, "FirstTable", "[#Headers],[ColB]"]' do
    @table.reference_for("FirstTable", "[#Headers],[ColB]","sheet1","F5").should == [:sheet_reference, "sheet1", [:cell, "C3"]]
  end

  it 'should be able to return a reference for [:table_reference, "FirstTable", "[#Totals],[ColA]"]' do
    @table.reference_for("FirstTable", "[#Totals],[ColA]","sheet1","F5").should == [:sheet_reference, "sheet1", [:cell, "B7"]]
  end
  
  it 'should be able to return a reference for [:table_reference, "FirstTable", "[#This Row],[ColA]"]' do
    @table.reference_for("FirstTable", "[#This Row],[ColA]","sheet1","F5").should == [:sheet_reference, "sheet1", [:cell, "B5"]]
  end
  
  it 'should be able to return a reference for [:table_reference, "FirstTable", "[#Headers],[ColB]:[ColC]"]' do
    @table.reference_for("FirstTable", "[#Headers],[ColB]:[ColC]","sheet1","F5").should == [:sheet_reference, "sheet1", [:area, "C3","D3"]]
  end

  it 'should be able to return a reference for [:table_reference, "FirstTable", "[#Totals],[ColA]:[ColC]"]' do
    @table.reference_for("FirstTable", "[#Totals],[ColA]:[ColC]","sheet1","F5").should == [:sheet_reference, "sheet1", [:area, "B7","D7"]]
  end

  it 'should be able to return a reference for [:table_reference, "FirstTable", "[#This Row],[ColB]:[ColC]"]' do
    @table.reference_for("FirstTable", "[#This Row],[ColB]:[ColC]","sheet1","F5").should == [:sheet_reference, "sheet1", [:area, "C5","D5"]]
  end    
end
