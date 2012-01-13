require_relative '../../spec_helper'

describe MapSheetNamesToRubyNames do
  
it "should take a file with worksheet names in its first column and create a file with worksheet names in the first column and a ruby equivalent in the second" do
input = <<END
sheet1	file1
A name with (unaceptable) characters	file2
A clashing name	file3
A (clashing) name	file4
A [clashing] name	file5
a [clashing]  name	file6
2010	file5
END

expected_output = <<END
sheet1	sheet1
A name with (unaceptable) characters	a_name_with_unaceptable_characters
A clashing name	a_clashing_name
A (clashing) name	a_clashing_name2
A [clashing] name	a_clashing_name3
a [clashing]  name	a_clashing_name4
2010	_2010
END

output = StringIO.new
MapSheetNamesToRubyNames.rewrite(input,output)
output.string.should == expected_output
end
end
