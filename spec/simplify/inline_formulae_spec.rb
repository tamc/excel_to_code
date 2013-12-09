require_relative '../spec_helper'

describe InlineFormulae do
  
it "should recursively work through formulae, inlining references where not important to function" do

input = <<END
A1\t[:cell, :"$A$2"]
A2\t[:cell, :"A3"]
A3\t[:number, 1]
A4\t[:sheet_reference,"sheet2",[:cell,:"A1"]]
A5\t[:sheet_reference,"sheet3",[:cell,:"A5"]]
A6\t[:function, "OFFSET", [:cell, :"$A$2"], [:cell, :"A3"], [:sheet_reference,:"sheet2",[:cell,:"A1"]]]
END

references = {
  [:sheet1, :A1] => [:cell, :"$A$2"],
  [:sheet1, :A2] => [:cell, :"A3"],
  [:sheet1, :A3] => [:number, 1],
  [:sheet2, :A1] => [:cell, :"A2"],
  [:sheet2, :A2] => [:sheet_reference,:sheet3,[:cell,:A1]],
  [:sheet3, :A1] => [:number, 5],
  [:sheet3, :A5] => [:number, 10]    
}

expected_output = <<END
A1\t[:number, 1]
A2\t[:number, 1]
A3\t[:number, 1]
A4\t[:number, 5]
A5\t[:number, 10]
A6\t[:function, "OFFSET", [:cell, :"$A$2"], [:number, 1], [:number, 5]]
END
    
input = StringIO.new(input)
output = StringIO.new
r = InlineFormulae.new
r.references = references
r.default_sheet_name = :sheet1
r.replace(input,output)
output.string.should == expected_output
end


it "should accept a block, which can be used to decide whether to inline a particualr reference or not" do

input = <<END
A1\t[:cell, :"$A$2"]
A2\t[:cell, :A3]
A3\t[:number, 1]
A4\t[:sheet_reference,:sheet2,[:cell,:A1]]
A5\t[:sheet_reference,:sheet3,[:cell,:A5]]
A6\t[:sheet_reference,:sheet1,[:cell, :"$A$2"]]
A7\t[:sheet_reference,:sheet2,[:cell, :"$A$2"]]
A8\t[:cell, :B8]
A9\t[:sheet_reference,:sheet2,[:cell, :"$B$8"]]
END

references = {
  [:sheet1, :A1] => [:cell, :"$A$2"],
  [:sheet1, :A2] => [:cell, :A3],
  [:sheet1, :A3] => [:number, 1],
  [:sheet2, :A1] => [:cell, :A2],
  [:sheet2, :A2] => [:sheet_reference,:sheet3,[:cell,:A1]],
  [:sheet3, :A1] => [:number, 5],
  [:sheet3, :A5] => [:number, 10]    
}

inline_ast_decision = lambda do |sheet,cell,references|
  if sheet == :sheet2 && cell == :A2
    false
  elsif sheet == :sheet3
    false
  else
    true
  end
end

expected_output = <<END
A1\t[:number, 1]
A2\t[:number, 1]
A3\t[:number, 1]
A4\t[:sheet_reference, :sheet2, [:cell, :A2]]
A5\t[:sheet_reference, :sheet3, [:cell, :A5]]
A6\t[:number, 1]
A7\t[:sheet_reference, :sheet2, [:cell, :"$A$2"]]
A8\t[:blank]
A9\t[:blank]
END
  
input = StringIO.new(input)
output = StringIO.new
r = InlineFormulae.new
r.references = references
r.default_sheet_name = :sheet1
r.inline_ast = inline_ast_decision
r.replace(input,output)
output.string.should == expected_output
end

end
