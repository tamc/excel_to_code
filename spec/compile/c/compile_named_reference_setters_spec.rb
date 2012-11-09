require_relative '../../spec_helper'

describe CompileNamedReferenceSetters do
  
def compile(named_references, sheet_names, cells_that_can_be_set_at_runtime)
  input = StringIO.new(named_references)
  sheet_names = StringIO.new(sheet_names)
  output = StringIO.new
  c = CompileNamedReferenceSetters.new
  c.cells_that_can_be_set_at_runtime = cells_that_can_be_set_at_runtime
  c.rewrite(named_references, sheet_names, output)
  output.string
end


it "should compile simple named references, skipping those that aren't settable" do

input = <<END
a	[:sheet_reference, "Sheet1", [:cell, "$B$1"]]
END

sheet_names = <<END
Sheet1\tsheet1
END

cells_that_can_be_set_at_runtime = {
  'Sheet1' => :all
}

expected = <<END
void set_a(ExcelValue newValue) {
  set_sheet1_b1(newValue);
}

END

compile(input, sheet_names, cells_that_can_be_set_at_runtime).should == expected

end

it "should compile named references that point to areas" do 
input = <<END
range	[:array, [:row, [:sheet_reference, "Sheet1", [:cell, "B1"]]], [:row, [:sheet_reference, "Sheet1", [:cell, "B2"]]], [:row, [:sheet_reference, "Sheet1", [:cell, "B3"]]]]
END

sheet_names = <<END
Sheet1\tsheet1
END

cells_that_can_be_set_at_runtime = {
  'Sheet1' => :all
}

expected = <<END
void set_range(ExcelValue newValue) {
  ExcelValue *array = newValue.array;
  set_sheet1_b1(array[0]);
  set_sheet1_b2(array[1]);
  set_sheet1_b3(array[2]);
}

END

compile(input, sheet_names, cells_that_can_be_set_at_runtime).should == expected

end

it "should not try and set a cell that cannot be set" do
input = <<END
range	[:array, [:row, [:sheet_reference, "Sheet1", [:cell, "B1"]]], [:row, [:sheet_reference, "Sheet1", [:cell, "B2"]]], [:row, [:sheet_reference, "Sheet1", [:cell, "B3"]]]]
END

sheet_names = <<END
Sheet1\tsheet1
END

cells_that_can_be_set_at_runtime = {
  'Sheet1' => ['B1']
}

expected = <<END
void set_range(ExcelValue newValue) {
  ExcelValue *array = newValue.array;
  set_sheet1_b1(array[0]);
  // sheet1_b2 not settable
  // sheet1_b3 not settable
}

END

compile(input, sheet_names, cells_that_can_be_set_at_runtime).should == expected

end

end
