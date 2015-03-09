require_relative '../spec_helper'

describe ExtractDataFromWorksheet do
  
  before do
    extractor = ExtractDataFromWorksheet.new
    input = excel_fragment 'FormulaeTypes.xml'
    extractor.extract(:SheetName, input)
    @extractor = extractor
  end

  it "should have a table_rids attribute with a hash keyed to the sheet name giving the ids of its tables" do
    extractor = ExtractDataFromWorksheet.new
    input = excel_fragment 'TableRelationships.xml'
    extractor.extract(:SheetName, input)
    extractor.table_rids.should == {:SheetName => ["rId1", "rId2", "rId3"] }
  end

  it "should have a worksheets_dimensions attribute with a hash keyed to the sheet name with the dimensions of the worksheet as the value" do
    extractor = ExtractDataFromWorksheet.new
    input = excel_fragment 'ValueTypes.xml'
    extractor.extract(:SheetName, input)
    extractor.worksheets_dimensions.should == {:SheetName => "A1:A6"}
  end

  it "should have a values attribute that returns a hash with the types and values of the cells" do
    extractor = ExtractDataFromWorksheet.new
    input = excel_fragment 'ValueTypes.xml'
    extractor.extract(:SheetName, input)
    extractor.values.should == {
      [:SheetName, :A1]  => [:boolean_true],
      [:SheetName, :A2]  => [:shared_string, 0],
      [:SheetName, :A3]  => [:number, 1],
      [:SheetName, :A4]  => [:number, 3.1415000000000002],
      [:SheetName, :A5]  => [:error, :'#NAME?'],
      [:SheetName, :A6]  => [:string, "Hello"],
    }
  end
    
  it "should have a formulae_simple attribute that returns a hash like [:SheetName, :A1] => [:arithmetic, [:number, 1], [:operator, :+], [:number, 1]]" do
    @extractor.formulae_simple.should == {
      [:SheetName, :B1] => [:arithmetic, [:number, 1.0], [:operator, :+], [:number, 1.0]],
      [:SheetName, :B2] => [:function, :COSH, [:arithmetic, [:number, 2.0], [:operator, :*], [:function, :PI]]]
    }
  end
  
  it "should have an formulae_array attribute that returns a Hash like ['SheetName', 'B3'] => ['B6:B8', [ast..]] for array formulae" do
    @extractor.formulae_array.should == {
      [:SheetName, :B5] => ["B5", [:area, :B1, :B4]],
      [:SheetName, :B6] => ["B6:B8", [:function, :IF, [:comparison, [:area, :B3, :B5], [:comparator, :"="], [:number, 8.0]], [:string, "Eight"], [:string, "Not Eight"]]],
    }
  end

  it "should have a formulae_shared attribute that returns a hash like [:SheetName, :B3] => ['B3:B4', '0', ast..]" do
    @extractor.formulae_shared.should == {
      [:SheetName, :B3] => ["B3:B4", "0", [:function, :COSH, [:arithmetic, [:number, 2.0], [:operator, :*], [:function, :PI]]]]
    }
  end

  it "should have a formulae_shared_targets attribute that returns a hash with ['sheetname', 'B1'] => '0' where '0' is the shared_target_number " do
    @extractor.formulae_shared_targets.should == {
      [:SheetName, :B3] => "0",
      [:SheetName, :B4] => "0",
    }
  end

  it "should have a :only_extract_values method which defaults to false" do
    extractor = ExtractDataFromWorksheet.new
    extractor.only_extract_values = true

    input = excel_fragment 'FormulaeTypes.xml'
    extractor.extract(:SheetName, input)
    extractor.formulae_simple.length.should == 0
    extractor.formulae_array.length.should == 0
    extractor.formulae_shared.length.should == 0
    extractor.formulae_shared_targets.length.should == 0
    extractor.values.length.should == 16
  end

  it "should convert Excels _x000D_ style escaping into proper unicode" do
    extractor = ExtractDataFromWorksheet.new
    extractor.convert_excels_unicode_escaping("One").should == "One"
    extractor.convert_excels_unicode_escaping("A_x000D_B_x000D_C").should == "A\rB\rC"
  end

end
