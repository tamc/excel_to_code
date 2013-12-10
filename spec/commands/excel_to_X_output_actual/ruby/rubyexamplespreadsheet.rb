# coding: utf-8
# Compiled version of /Users/tamc/Documents/github/excel_to_code/spec/test_data/ExampleSpreadsheet.xlsx
require '/Users/tamc/Documents/github/excel_to_code/src/excel/excel_functions'

class RubyExampleSpreadsheet
  include ExcelFunctions
  def original_excel_filename; "/Users/tamc/Documents/github/excel_to_code/spec/test_data/ExampleSpreadsheet.xlsx"; end
  attr_accessor :valuetypes_a1 # Default: true
  attr_accessor :valuetypes_a2 # Default: "Hello"
  attr_accessor :valuetypes_a3 # Default: 1.0
  attr_accessor :valuetypes_a4 # Default: 3.1415
  attr_accessor :valuetypes_a5 # Default: :name
  attr_accessor :valuetypes_a6 # Default: "Hello"
  def formulaetypes_a1; @formulaetypes_a1 ||= "Simple"; end
  def formulaetypes_b1; @formulaetypes_b1 ||= 2.0; end
  def formulaetypes_a2; @formulaetypes_a2 ||= "Sharing"; end
  def formulaetypes_b2; @formulaetypes_b2 ||= 267.7467614837482; end
  def formulaetypes_a3; @formulaetypes_a3 ||= "Shared"; end
  def formulaetypes_b3; @formulaetypes_b3 ||= 267.7467614837482; end
  def formulaetypes_a4; @formulaetypes_a4 ||= "Shared"; end
  def formulaetypes_b4; @formulaetypes_b4 ||= 267.7467614837482; end
  def formulaetypes_a5; @formulaetypes_a5 ||= "Array (single)"; end
  def formulaetypes_b5; @formulaetypes_b5 ||= 2.0; end
  def formulaetypes_a6; @formulaetypes_a6 ||= "Arraying (multiple)"; end
  def formulaetypes_b6; @formulaetypes_b6 ||= "Not Eight"; end
  def formulaetypes_a7; @formulaetypes_a7 ||= "Arrayed (multiple)"; end
  def formulaetypes_b7; @formulaetypes_b7 ||= "Not Eight"; end
  def formulaetypes_a8; @formulaetypes_a8 ||= "Arrayed (multiple)"; end
  def formulaetypes_b8; @formulaetypes_b8 ||= "Not Eight"; end
  def ranges_b1; @ranges_b1 ||= "This sheet"; end
  def ranges_c1; @ranges_c1 ||= "Other sheet"; end
  def ranges_a2; @ranges_a2 ||= "Standard"; end
  def ranges_b2; @ranges_b2 ||= sum([[ranges_f4],[ranges_f5],[ranges_f6]]); end
  def ranges_c2; @ranges_c2 ||= sum([[valuetypes_a3],[valuetypes_a4]]); end
  def ranges_a3; @ranges_a3 ||= "Column"; end
  def ranges_b3; @ranges_b3 ||= sum([[nil],[nil],[nil],[ranges_f4],[ranges_f5],[ranges_f6]]); end
  def ranges_c3; @ranges_c3 ||= sum([[valuetypes_a1],[valuetypes_a2],[valuetypes_a3],[valuetypes_a4],[valuetypes_a5],[valuetypes_a6]]); end
  def ranges_a4; @ranges_a4 ||= "Row"; end
  def ranges_b4; @ranges_b4 ||= sum([[nil,nil,nil,nil,ranges_e5,ranges_f5,ranges_g5]]); end
  def ranges_c4; @ranges_c4 ||= sum(valuetypes_a4); end
  attr_accessor :ranges_f4 # Default: 1.0
  attr_accessor :ranges_e5 # Default: 1.0
  attr_accessor :ranges_f5 # Default: 2.0
  attr_accessor :ranges_g5 # Default: 3.0
  attr_accessor :ranges_f6 # Default: 3.0
  def referencing_a1; @referencing_a1 ||= referencing_c4; end
  def referencing_a2; @referencing_a2 ||= referencing_c4; end
  attr_accessor :referencing_a4 # Default: 10.0
  def referencing_b4; @referencing_b4 ||= common0; end
  def referencing_c4; @referencing_c4 ||= add(common0,1.0); end
  def referencing_a5; @referencing_a5 ||= 3.0; end
  def referencing_b8; @referencing_b8 ||= referencing_c4; end
  def referencing_b9; @referencing_b9 ||= common1; end
  def referencing_b11; @referencing_b11 ||= "Named"; end
  def referencing_c11; @referencing_c11 ||= "Reference"; end
  def tables_a1; @tables_a1 ||= nil; end
  attr_accessor :tables_b2 # Default: "ColA"
  attr_accessor :tables_c2 # Default: "ColB"
  attr_accessor :tables_d2 # Default: "Column1"
  attr_accessor :tables_b3 # Default: 1.0
  attr_accessor :tables_c3 # Default: "A"
  def tables_d3; @tables_d3 ||= string_join(tables_b3,tables_c3); end
  attr_accessor :tables_b4 # Default: 2.0
  attr_accessor :tables_c4 # Default: "B"
  def tables_d4; @tables_d4 ||= string_join(tables_b4,tables_c4); end
  def tables_f4; @tables_f4 ||= tables_c4; end
  def tables_g4; @tables_g4 ||= excel_match("2B",[[tables_b4,tables_c4,tables_d4]],false); end
  def tables_h4; @tables_h4 ||= excel_match("B",[[tables_c4,tables_d4]]); end
  def tables_b5; @tables_b5 ||= common3; end
  def tables_c5; @tables_c5 ||= sum([[tables_c3],[tables_c4]]); end
  def tables_e6; @tables_e6 ||= tables_b2; end
  def tables_f6; @tables_f6 ||= tables_c2; end
  def tables_g6; @tables_g6 ||= tables_d2; end
  def tables_e7; @tables_e7 ||= tables_b5; end
  def tables_f7; @tables_f7 ||= tables_c5; end
  def tables_g7; @tables_g7 ||= nil; end
  def tables_e8; @tables_e8 ||= tables_b2; end
  def tables_f8; @tables_f8 ||= tables_c2; end
  def tables_g8; @tables_g8 ||= tables_d2; end
  def tables_e9; @tables_e9 ||= tables_b3; end
  def tables_f9; @tables_f9 ||= tables_c3; end
  def tables_g9; @tables_g9 ||= tables_d3; end
  def tables_c10; @tables_c10 ||= common1; end
  def tables_e10; @tables_e10 ||= tables_b4; end
  def tables_f10; @tables_f10 ||= tables_c4; end
  def tables_g10; @tables_g10 ||= tables_d4; end
  def tables_c11; @tables_c11 ||= common3; end
  def tables_e11; @tables_e11 ||= tables_b5; end
  def tables_f11; @tables_f11 ||= tables_c5; end
  def tables_g11; @tables_g11 ||= nil; end
  def tables_c12; @tables_c12 ||= tables_b5; end
  def tables_c13; @tables_c13 ||= common5; end
  def tables_c14; @tables_c14 ||= common5; end
  def s_innapropriate_sheet_name__c4; @s_innapropriate_sheet_name__c4 ||= valuetypes_a3; end
  def common0; @common0 ||= add(referencing_a4,1.0); end
  def common1; @common1 ||= sum([[tables_b5,tables_c5,nil]]); end
  def common2; @common2 ||= [[tables_b5,tables_c5,nil]]; end
  def common3; @common3 ||= sum([[tables_b3],[tables_b4]]); end
  def common4; @common4 ||= [[tables_b3],[tables_b4]]; end
  def common5; @common5 ||= sum([[tables_b3,tables_c3,tables_d3],[tables_b4,tables_c4,tables_d4]]); end
  def common6; @common6 ||= [[tables_b3,tables_c3,tables_d3],[tables_b4,tables_c4,tables_d4]]; end


  # starting initializer
  def initialize
    @valuetypes_a1 = true
    @valuetypes_a2 = "Hello"
    @valuetypes_a3 = 1.0
    @valuetypes_a4 = 3.1415
    @valuetypes_a5 = :name
    @valuetypes_a6 = "Hello"
    @ranges_f4 = 1.0
    @ranges_e5 = 1.0
    @ranges_f5 = 2.0
    @ranges_g5 = 3.0
    @ranges_f6 = 3.0
    @referencing_a4 = 10.0
    @tables_b2 = "ColA"
    @tables_c2 = "ColB"
    @tables_d2 = "Column1"
    @tables_b3 = 1.0
    @tables_c3 = "A"
    @tables_b4 = 2.0
    @tables_c4 = "B"
  end

end
