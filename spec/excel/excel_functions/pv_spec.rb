require_relative '../../spec_helper.rb'

describe "ExcelFunctions: PV(rate, nper, pmt, fv, type)" do
  
  it "should return the present value of an investment" do
    FunctionTest.pv(0.03,12.0,100).round.should == -995
    FunctionTest.pv(0.03,12.0,-100.0,100.0).round.should == 925
    FunctionTest.pv(0.03,12.0,-100,-100,1).round.should == 1095
    FunctionTest.pv(0.03,12.0,nil,100).round.should == -70
    FunctionTest.pv(0.03,nil,100,100).round.should == -100
    FunctionTest.pv(nil,12,100,100).round.should == -1300
    FunctionTest.pv(0.03,12,nil,nil,nil).round.should == 0
  end

  it "should return an error when given inappropriate arguments" do
    FunctionTest.pv("Asdasddf",12,0).should == :value
  end
    
  it "should return an error if an argument is an error" do
    FunctionTest.pv(:error,12,100).should == :error
    FunctionTest.pv(0.03,:error,100).should == :error
    FunctionTest.pv(0.03,12,:error).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'PV'].should == 'pv'
  end
  
end
