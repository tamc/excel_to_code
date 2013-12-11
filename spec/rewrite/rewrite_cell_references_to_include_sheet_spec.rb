require_relative '../spec_helper'

describe RewriteCellReferencesToIncludeSheet do
  
it "should take a file with ast in its last column and rewrite any [:cell, 'A1'] and [:area, 'A1', 'B2'] references to [:sheet_reference, 'sheet1', [:cell, 'A1']]" do
input = <<END
B2	[:function, "SUM", [:cell, "F4"]]
C2	[:function, "SUM", [:sheet_reference, "ValueTypes", [:area, "A3", "A4"]]]
B3	[:function, "SUM", [:area, "F1", "F6"]]
END

input = StringIO.new(input)
output = StringIO.new

r = RewriteCellReferencesToIncludeSheet.new
r.worksheet = "default"
r.rewrite(input,output)

expected =<<END
B2	[:function, "SUM", [:sheet_reference, :default, [:cell, :F4]]]
C2	[:function, "SUM", [:sheet_reference, :ValueTypes, [:area, :A3, :A4]]]
B3	[:function, "SUM", [:sheet_reference, :default, [:area, :F1, :F6]]]
END
output.string.should == expected

end

# Can't get this to work at the moment
# it "should return the same object when referencing the same worksheet" do 
#   r = RewriteCellReferencesToIncludeSheetAst.new
#   r.worksheet = :sheet1
#   first = r.map([:cell, :A1])
#   second = r.map([:cell, :A1])
#   first.object_id.should == second.object_id
# end

# Can't get this to work at the moment
#it "should ensure that sheet references are also the same object" do
#  r = RewriteCellReferencesToIncludeSheetAst.new
#  r.worksheet = :sheet1
#  first = r.map([:sheet_reference, :shee2, [:cell, :A1]])
#  second = r.map([:sheet_reference, :shee2, [:cell, :A1]])
#  first.object_id.should == second.object_id
#end

it "should remove fixed (e.g., $A$1) references" do
  r = RewriteCellReferencesToIncludeSheetAst.new
  r.worksheet = :sheet1
  r.map([:sheet_reference, :shee2, [:cell, :"$A$1"]]).should == [:sheet_reference, :shee2, [:cell, :A1]]
  r.map([:cell, :"$A$1"]).should == [:sheet_reference, :sheet1, [:cell, :A1]]
  r.map([:area, :"$A$1", :"B$10"]).should == [:sheet_reference, :sheet1, [:area, :A1, :B10]]
end

end
