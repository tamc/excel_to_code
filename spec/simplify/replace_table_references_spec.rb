require_relative '../spec_helper'

describe ReplaceTableReferences do
  
it "should replace table references with cell and array references" do

input = <<END
A3\t[:table_reference, "FirstTable", "[#This Row],[ColA]:[ColB]"]
A4\t[:string_join, [:table_reference, "FirstTable", "[#This Row],[ColA]"], [:table_reference, "FirstTable", "[#This Row],[ColB]"]]
A5\t[:function,"SUM",[:cell,"A1"],[:table_reference, "FirstTable", "[#This Row],[ColA]:[ColB]"]]
C2\t[:table_reference, "FirstTable", "ColA"]
END

tables = <<END
FirstTable	Tables	B2:C5	1	ColA	ColB
END

expected_output = <<END
A3\t[:sheet_reference, "Tables", [:area, "B3", "C3"]]
A4\t[:string_join, [:sheet_reference, "Tables", [:cell, "B4"]], [:sheet_reference, "Tables", [:cell, "C4"]]]
A5\t[:function, "SUM", [:cell, "A1"], [:sheet_reference, "Tables", [:area, "B5", "C5"]]]
C2\t[:sheet_reference, "Tables", [:cell, "B2"]]
END
    
input = StringIO.new(input)
tables = StringIO.new(tables)
output = StringIO.new
r = ReplaceTableReferences.new
r.sheet_name = "Tables"
r.replace(input,tables,output)
output.string.should == expected_output
end # /it

it "should work when applied to array references" do

input = <<END
A3\tA3:C3\t[:table_reference, "FirstTable", "[#This Row],[ColA]:[ColB]"]
A4\tA4:C4\t[:string_join, [:table_reference, "FirstTable", "[#This Row],[ColA]"], [:table_reference, "FirstTable", "[#This Row],[ColB]"]]
END

tables = <<END
FirstTable	Tables	B2:C5	1	ColA	ColB
END

expected_output = <<END
A3\tA3:C3\t[:sheet_reference, "Tables", [:area, "B3", "C3"]]
A4\tA4:C4\t[:string_join, [:sheet_reference, "Tables", [:cell, "B4"]], [:sheet_reference, "Tables", [:cell, "C4"]]]
END
    
input = StringIO.new(input)
tables = StringIO.new(tables)
output = StringIO.new
r = ReplaceTableReferences.new
r.sheet_name = "Tables"
r.replace(input,tables,output)
output.string.should == expected_output
end # /it


end # /describe
