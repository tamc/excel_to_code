require_relative '../../spec_helper.rb'
include ExcelFunctions

describe "ExcelFunctions: excel_match(lookup_value,lookup_array,match_type)" do
  
  it "should return the first matching value in the lookup array, or :na if not found" do
    excel_match(10.0,[[10,100]],0.0).should == 1
    excel_match(100.0,[[10,100]],0.0).should == 2
    excel_match(1000.0,[[10,100]],0.0).should == :na
    excel_match('bEAr',[["Pear"],["Bear"],["Apple"]],0.0).should == 2
    excel_match(1000.0,[[10,100]],1.0).should == 2
    excel_match(1.0,[[10,100]],1.0).should == :na
    excel_match('Care',[["Pear"],["Bear"],["Apple"]],-1.0).should == 1  
    excel_match('Zebra',[["Pear"],["Bear"],["Apple"]],-1.0).should == :na
    excel_match('a',[["Pear"],["Bear"],["Apple"]],-1.0).should == 2
  end
  
  it "should treat a single cell given as a lookup_array as an array of size 1" do
    excel_match(10.0,10,0.0).should == 1
    excel_match(20.0,10,0.0).should == :na
  end
  
  it "should treat nil as zero" do
      excel_match(0,[1,nil,0]).should == 2
      excel_match(nil,[1,0,nil]).should == 2
      excel_match(10.0,[10,100],nil).should == 1
    end

   it "should return an error if any argument is an error" do
     excel_match(:error,[1],0).should == :error
     excel_match(1,:error,0).should == :error
     excel_match(1,[1],:error).should == :error
     excel_match(:error1,:error2,:error3).should == :error1
   end
  
end


