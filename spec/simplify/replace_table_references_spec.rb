require_relative '../spec_helper'

describe ReplaceTableReferences do
  
it "should replace table references with cell and array references" do

input = <<END
A3\t[:table_reference, "FirstTable", "[#This Row],[ColA]:[ColB]"]
A4\t[:string_join, [:table_reference, "FirstTable", "[#This Row],[ColA]"], [:table_reference, "FirstTable", "[#This Row],[ColB]"]]
A5\t[:function,"SUM",[:cell,"A1"],[:table_reference, "FirstTable", "[#This Row],[ColA]:[ColB]"]]
C3\t[:table_reference, "MissingTable", "ColA"]
C4\t[:table_reference, "FirstTable", "ColA"]
END

tables = <<END
FirstTable	Tables	B2:C5	1	ColA	ColB
END

expected_output = <<END
A3\t[:sheet_reference, "Tables", [:area, "B3", "C3"]]
A4\t[:string_join, [:sheet_reference, "Tables", [:cell, "B4"]], [:sheet_reference, "Tables", [:cell, "C4"]]]
A5\t[:function, "SUM", [:cell, "A1"], [:sheet_reference, "Tables", [:area, "B5", "C5"]]]
C3\t[:error, "#REF!"]
C4\t[:sheet_reference, "Tables", [:cell, "B4"]]
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

it "should work when applied to array references" do

input = <<END
F15\t[:table_reference, "I.a.Inputs", "#Headers"]
F15\t[:function, "INDEX", [:function, "INDIRECT", [:string_join, [:string, "'"], [:table_reference, "I.a.Inputs", "#Headers"], [:string, "'!Year.Matrix"]]], [:function, "MATCH", [:string_join, [:string, "Subtotal."], [:cell, "$A$2"]], [:function, "INDIRECT", [:string_join, [:string, "'"], [:table_reference, "I.a.Inputs", "#Headers"], [:string, "'!Year.Modules"]]], [:number, "0"]], [:function, "MATCH", [:table_reference, "I.a.Inputs", "Vector"], [:function, "INDIRECT", [:string_join, [:string, "'"], [:table_reference, "I.a.Inputs", "#Headers"], [:string, "'!Year.Vectors"]]], [:number, "0"]]]
END

tables = <<END
I.a.Inputs	I.a	C14:O16	0	Vector	Name	Notes	2007	2010	2015	2020	2025	2030	2035	2040	2045	2050
END

expected_output = <<END
F15\t[:sheet_reference, "I.a", [:cell, "F14"]]
F15\t[:function, "INDEX", [:function, "INDIRECT", [:string_join, [:string, "'"], [:sheet_reference, "I.a", [:cell, "F14"]], [:string, "'!Year.Matrix"]]], [:function, "MATCH", [:string_join, [:string, "Subtotal."], [:cell, "$A$2"]], [:function, "INDIRECT", [:string_join, [:string, "'"], [:sheet_reference, "I.a", [:cell, "F14"]], [:string, "'!Year.Modules"]]], [:number, "0"]], [:function, "MATCH", [:sheet_reference, "I.a", [:cell, "C15"]], [:function, "INDIRECT", [:string_join, [:string, "'"], [:sheet_reference, "I.a", [:cell, "F14"]], [:string, "'!Year.Vectors"]]], [:number, "0"]]]
END
    
input = StringIO.new(input)
tables = StringIO.new(tables)
output = StringIO.new
r = ReplaceTableReferences.new
r.sheet_name = "I.a"
r.replace(input,tables,output)
output.string.should == expected_output
end # /it

end # /describe
