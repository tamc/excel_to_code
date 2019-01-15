require_relative '../../spec_helper.rb'

describe "ExcelFunctions: CEILING" do
  

    # Number\tSignificance\tMode\tCEILING.MATH(number, significance, mode)
    tests = <<-END
0.1	0.1	0	0.1
-0.1	0.1	0	-0.1
-0.1	0.1	1	-0.1
136	10	0	140
-136	10	0	-130
-136	10	1	-140    
END
  tests = tests.split("\n").map { |l| l.split("\t").map(&:strip) }
  tests.each do |t|
    expected = t.pop.to_f
      it "#{t}" do
        FunctionTest.ceiling(*t).should == expected
      end
  end

  it "should return an error when given inappropriate arguments" do
    FunctionTest.ceiling("Asdasddf",1,1).should == :value
  end
    
  it "should return an error if an argument is an error" do
    FunctionTest.ceiling(:error,1,1).should == :error
  end
  
  it "should be in the list of functions that can be mapped to ruby" do
    MapFormulaeToRuby::FUNCTIONS[:'CEILING'].should == 'ceiling'
  end
  
end
