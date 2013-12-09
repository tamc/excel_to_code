require_relative '../../spec_helper.rb'

describe "ExcelFunctions: PMT(rate,number_of_periods,present_value) - optional arguments not yet implemented" do
  
  it "should calculate the monthly payment required for a given principal, interest rate and loan period" do
    FunctionTest.pmt(0.1,10,100).should be_within(0.01).of(-16.27)
    FunctionTest.pmt(0.0123,99.1,123.32).should be_within(0.01).of(-2.159)
    FunctionTest.pmt(0,2,10).should be_within(0.01).of(-5)
  end

  it "should work if arguments are given as strings, so long as the strings contain numbers" do
    FunctionTest.pmt('0.1','10','100').should be_within(0.01).of(-16.27)
  end
  
  it "should work if arguments given as booleans, with true = 1 and false = 0" do
    FunctionTest.pmt(false,true,true).should be_within(0.01).of(-1)
  end
    
  it "should treat nil as zero" do
    FunctionTest.pmt(nil,1,nil).should == 0
  end
  
  it "should return an error if an argument is an error" do
    FunctionTest.pmt(:error1,10,100).should == :error1
    FunctionTest.pmt(0.1,:error2,100).should == :error2
    FunctionTest.pmt(0.1,10,:error3).should == :error3
    FunctionTest.pmt(:error1,:error2,:error3).should == :error1
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'PMT'].should == 'pmt'
  end
  
end
