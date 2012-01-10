require_relative '../spec_helper'

describe ReplaceTableReferences do
  
  it "should replace table references with cell and array references" do

input = <<END
A3\t[:table_reference, "FirstTable", "[#This Row],[ColA]:[ColB]"]
END

tables = <<END
FirstTable	Tables	B2:C5	1	ColA	ColB
END

expected_output = <<END
A3\t[:sheet_reference, "Tables", [:area, "B3", "C3"]]
END
    
    input = StringIO.new(input)
    tables = StringIO.new(tables)
    output = StringIO.new
    ReplaceTableReferences.replace(input,"Tables",tables,output)
    output.string.should == expected_output
  end

end
