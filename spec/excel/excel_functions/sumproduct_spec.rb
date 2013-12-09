require_relative '../../spec_helper.rb'

describe "ExcelFunctions: SUMPRODUCT" do
  
  it "should multiply together and then sum the elements in row or column areas given as arguments" do
    FunctionTest.sumproduct([[10],[100],[nil]],[[nil],[100],[10]]).should == 100*100
  end

  it "should return :value when miss-matched array sizes" do
    FunctionTest.sumproduct([[10],[100],[nil]],[[nil]]).should == :value
  end

  it "if all its arguments are single values, should multiply them together" do
    FunctionTest.sumproduct(10,100,1000).should == 10*100*1000
  end
  
  it "if it only has one range as an argument, should add its elements together" do
    FunctionTest.sumproduct([[1],[2],[3]]).should == 1 + 2 + 3
  end
  
  it "if given multi row and column areas as arguments, should multipy the corresponding cell in each area and then add them all" do
    FunctionTest.sumproduct([[1,2],[4,5]],[[10,20],[40,50]],[[11,21],[41,51]]).should == 1*10*11 + 2*20*21 + 4*40*41 + 5*50*51
  end
    
  it "should raise an error if nil values outside of an array" do
    FunctionTest.sumproduct(nil,1).should == :value
  end

  it "should ignore non-numeric values within an array" do
    FunctionTest.sumproduct([[nil],[nil]],[[nil],[nil]]).should == 0
  end

  it "should return an error if an argument is an error" do
    FunctionTest.sumproduct(:error).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'SUMPRODUCT'].should == 'sumproduct'
  end
  
end
