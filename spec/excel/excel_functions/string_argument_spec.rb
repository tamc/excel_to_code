require_relative '../../spec_helper.rb'

describe "ExcelFunctions: string_argument(thing)" do
  
  {
    "string" => "string",
    nil => "",
    1 => "1",
    1.23 => "1.23",
    true => "TRUE",
    false => "FALSE",
    :name => "#NAME?",
    :value => "#VALUE!",
    :div0 => "#DIV/0!",
    :ref => "#REF!",
    :na => "#N/A",
    :num => "#NUM!",
  }.each do |value, string|
    it "string_argument(#{value}) == #{string.inspect}" do
      FunctionTest.string_argument(value).should == string
    end
  end
end
