require_relative '../spec_helper'

describe ReplaceTransposeFunction do

  it "should replace TRANSPOSE() functions with transposed arrays" do

    r = ReplaceTransposeFunction.new
    r.replace([:function, :TRANSPOSE, [:array, [:row, [:number, 1], [:number, 2]], [:row, [:number, 3], [:number, 4]]]]).should == [:array, [:row, [:number, 1], [:number, 3]], [:row, [:number, 2], [:number, 4]]]
    r.replace([:array, [:row, [:number, 1], [:number, 2]], [:row, [:number, 3], [:number, 4]]]).should == [:array, [:row, [:number, 1], [:number, 2]], [:row, [:number, 3], [:number, 4]]]
  end


end
