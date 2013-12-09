require File.expand_path(File.dirname(__FILE__) + '/../spec_helper')

describe AstCopyFormula do
  
  it "should take a specified distance in rows and columns, and move all references by the appropriate ammount" do
    copy = AstCopyFormula.new
    copy.rows_to_move = 1
    copy.columns_to_move = 1
    copy.copy([:cell,:A1]).should == [:cell,:B2]
    copy.copy([:cell,:'$A1']).should == [:cell,:'$A2']
    copy.copy([:cell,:'A$1']).should == [:cell,:'B$1']
    copy.copy([:cell,:'$A$1']).should == [:cell,:'$A$1']
    copy.copy([:area,:'A1',:'Z10']).should == [:area,:'B2',:'AA11']
    copy.copy([:area,:'$A$1',:'Z$10']).should == [:area,:'$A$1',:'AA$10']
    copy.copy([:sheet_reference,:sheet1,[:cell,:A1]]).should == [:sheet_reference,:sheet1,[:cell,:B2]]
  end
  
  it "should throw an exception if column or row ranges are in the AST" do
    copy = AstCopyFormula.new
    lambda { copy.copy([:column_range,'C:F'])}.should raise_exception(NotSupportedException)
    lambda { copy.copy([:row_range,'10:12'])}.should raise_exception(NotSupportedException)
  end

  it "should not duplicate constant values" do
    copy = AstCopyFormula.new
    a = [:artithmetic, [:number, 2], [:operator, "+"], [:number, 3]]
    b = copy.copy(a)
    a[1].object_id.should == b[1].object_id
    a[2].object_id.should == b[2].object_id
    a[3].object_id.should == b[3].object_id
  end
  
end
