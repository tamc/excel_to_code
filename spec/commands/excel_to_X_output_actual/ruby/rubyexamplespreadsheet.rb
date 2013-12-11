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
  attr_accessor :referencing_c15 # Default: 1.0
  attr_accessor :referencing_d15 # Default: 2.0
  attr_accessor :referencing_e15 # Default: 3.0
  attr_accessor :referencing_f15 # Default: 4.0
  attr_accessor :referencing_c16 # Default: 1.4535833325868115
  attr_accessor :referencing_d16 # Default: 1.4535833325868115
  attr_accessor :referencing_e16 # Default: 1.511726665890284
  attr_accessor :referencing_f16 # Default: 1.5407983325420203
  attr_accessor :referencing_c17 # Default: 9.054545454545455
  attr_accessor :referencing_d17 # Default: 12.0
  attr_accessor :referencing_e17 # Default: 18.0
  attr_accessor :referencing_f17 # Default: 18.0
  attr_accessor :referencing_c18 # Default: 0.3681150635671386
  attr_accessor :referencing_d18 # Default: 0.3681150635671386
  attr_accessor :referencing_e18 # Default: 0.40588480110308967
  attr_accessor :referencing_f18 # Default: 0.42190146532760275
  attr_accessor :referencing_c19 # Default: 0.651
  attr_accessor :referencing_d19 # Default: 0.651
  attr_accessor :referencing_e19 # Default: 0.651
  attr_accessor :referencing_f19 # Default: 0.651
  attr_accessor :referencing_c22 # Default: 4.0
  def referencing_d22; @referencing_d22 ||= index(common3,1.0,1.0); end
  def referencing_d23; @referencing_d23 ||= index(common3,2.0,1.0); end
  def referencing_d24; @referencing_d24 ||= index(common3,3.0,1.0); end
  def referencing_d25; @referencing_d25 ||= index(common3,4.0,1.0); end
  def referencing_c31; @referencing_c31 ||= "Technology efficiencies -- hot water -- annual mean"; end
  def referencing_o31; @referencing_o31 ||= "% of input energy"; end
  def referencing_f33; @referencing_f33 ||= "Electricity (delivered to end user)"; end
  def referencing_g33; @referencing_g33 ||= "Electricity (supplied to grid)"; end
  def referencing_h33; @referencing_h33 ||= "Solid hydrocarbons"; end
  def referencing_i33; @referencing_i33 ||= "Liquid hydrocarbons"; end
  def referencing_j33; @referencing_j33 ||= "Gaseous hydrocarbons"; end
  def referencing_k33; @referencing_k33 ||= "Heat transport"; end
  def referencing_l33; @referencing_l33 ||= "Environmental heat"; end
  def referencing_m33; @referencing_m33 ||= "Heating & cooling"; end
  def referencing_n33; @referencing_n33 ||= "Conversion losses"; end
  def referencing_o33; @referencing_o33 ||= "Balance"; end
  def referencing_c34; @referencing_c34 ||= "Code"; end
  def referencing_d34; @referencing_d34 ||= "Technology"; end
  def referencing_e34; @referencing_e34 ||= "Notes"; end
  attr_accessor :referencing_f34 # Default: "V.01"
  attr_accessor :referencing_g34 # Default: "V.02"
  attr_accessor :referencing_h34 # Default: "V.03"
  attr_accessor :referencing_i34 # Default: "V.04"
  attr_accessor :referencing_j34 # Default: "V.05"
  attr_accessor :referencing_k34 # Default: "V.07"
  attr_accessor :referencing_l34 # Default: "R.07"
  attr_accessor :referencing_m34 # Default: "H.01"
  attr_accessor :referencing_n34 # Default: "X.01"
  def referencing_c35; @referencing_c35 ||= 1.0; end
  def referencing_d35; @referencing_d35 ||= "Gas boiler (old)"; end
  attr_accessor :referencing_j35 # Default: -1.0
  attr_accessor :referencing_m35 # Default: 0.76
  attr_accessor :referencing_n35 # Default: 0.24
  def referencing_o35; @referencing_o35 ||= 0.0; end
  def referencing_c36; @referencing_c36 ||= 2.0; end
  def referencing_d36; @referencing_d36 ||= "Gas boiler (new)"; end
  attr_accessor :referencing_j36 # Default: -1.0
  attr_accessor :referencing_m36 # Default: 0.91
  attr_accessor :referencing_n36 # Default: 0.09
  def referencing_o36; @referencing_o36 ||= 0.0; end
  def referencing_c37; @referencing_c37 ||= 3.0; end
  def referencing_d37; @referencing_d37 ||= "Resistive heating"; end
  attr_accessor :referencing_f37 # Default: -1.0
  attr_accessor :referencing_m37 # Default: 1.0
  attr_accessor :referencing_n37 # Default: 0.0
  def referencing_o37; @referencing_o37 ||= 0.0; end
  def referencing_c38; @referencing_c38 ||= 4.0; end
  def referencing_d38; @referencing_d38 ||= "Oil-fired boiler"; end
  attr_accessor :referencing_i38 # Default: -1.0
  attr_accessor :referencing_m38 # Default: 0.97
  attr_accessor :referencing_n38 # Default: 0.03
  def referencing_o38; @referencing_o38 ||= -2.7755575615628914e-17; end
  def referencing_c39; @referencing_c39 ||= 5.0; end
  def referencing_d39; @referencing_d39 ||= "Solid-fuel boiler"; end
  def referencing_e39; @referencing_e39 ||= "[2]"; end
  attr_accessor :referencing_h39 # Default: -1.0
  attr_accessor :referencing_m39 # Default: 0.87
  attr_accessor :referencing_n39 # Default: 0.13
  def referencing_o39; @referencing_o39 ||= 0.0; end
  def referencing_c40; @referencing_c40 ||= 6.0; end
  def referencing_d40; @referencing_d40 ||= "Stirling engine micro-CHP"; end
  def referencing_e40; @referencing_e40 ||= "[3]"; end
  attr_accessor :referencing_g40 # Default: 0.225
  attr_accessor :referencing_j40 # Default: -1.0
  attr_accessor :referencing_m40 # Default: 0.63
  attr_accessor :referencing_n40 # Default: 0.145
  def referencing_o40; @referencing_o40 ||= 0.0; end
  def referencing_c41; @referencing_c41 ||= 7.0; end
  def referencing_d41; @referencing_d41 ||= "Fuel-cell micro-CHP"; end
  def referencing_e41; @referencing_e41 ||= "[3]"; end
  attr_accessor :referencing_g41 # Default: 0.45
  attr_accessor :referencing_j41 # Default: -1.0
  attr_accessor :referencing_m41 # Default: 0.45
  attr_accessor :referencing_n41 # Default: 0.1
  def referencing_o41; @referencing_o41 ||= 0.0; end
  def referencing_c42; @referencing_c42 ||= 8.0; end
  def referencing_d42; @referencing_d42 ||= "Air-source heat pump"; end
  attr_accessor :referencing_f42 # Default: -1.0
  attr_accessor :referencing_l42 # Default: -1.0
  attr_accessor :referencing_m42 # Default: 2.0
  def referencing_o42; @referencing_o42 ||= 0.0; end
  def referencing_c43; @referencing_c43 ||= 9.0; end
  def referencing_d43; @referencing_d43 ||= "Ground-source heat pump"; end
  attr_accessor :referencing_f43 # Default: -1.0
  attr_accessor :referencing_l43 # Default: -2.0
  attr_accessor :referencing_m43 # Default: 3.0
  def referencing_o43; @referencing_o43 ||= 0.0; end
  def referencing_c44; @referencing_c44 ||= 10.0; end
  def referencing_d44; @referencing_d44 ||= "Geothermal electricity"; end
  attr_accessor :referencing_l44 # Default: -1.0
  attr_accessor :referencing_m44 # Default: 0.85
  attr_accessor :referencing_n44 # Default: 0.15
  def referencing_o44; @referencing_o44 ||= 0.0; end
  def referencing_c45; @referencing_c45 ||= 11.0; end
  def referencing_d45; @referencing_d45 ||= "Community scale gas CHP with local district heating"; end
  attr_accessor :referencing_g45 # Default: 0.38
  attr_accessor :referencing_j45 # Default: -1.0
  attr_accessor :referencing_m45 # Default: 0.38
  attr_accessor :referencing_n45 # Default: 0.24
  def referencing_o45; @referencing_o45 ||= 0.0; end
  def referencing_c46; @referencing_c46 ||= 12.0; end
  def referencing_d46; @referencing_d46 ||= "Community scale solid-fuel CHP with local district heating"; end
  attr_accessor :referencing_g46 # Default: 0.17
  attr_accessor :referencing_h46 # Default: -1.0
  attr_accessor :referencing_m46 # Default: 0.57
  attr_accessor :referencing_n46 # Default: 0.26
  def referencing_o46; @referencing_o46 ||= 0.0; end
  def referencing_c47; @referencing_c47 ||= 13.0; end
  def referencing_d47; @referencing_d47 ||= "Long distance district heating from large power stations"; end
  def referencing_e47; @referencing_e47 ||= "[6]"; end
  attr_accessor :referencing_k47 # Default: -1.0
  attr_accessor :referencing_m47 # Default: 0.9
  attr_accessor :referencing_n47 # Default: 0.1
  def referencing_o47; @referencing_o47 ||= 0.0; end
  def referencing_d50; @referencing_d50 ||= "Gas boiler (old)"; end
  attr_accessor :referencing_g50 # Default: 137.26515207025273
  def referencing_d51; @referencing_d51 ||= "Gas boiler (new)"; end
  attr_accessor :referencing_g51 # Default: 30.731004194832696
  def referencing_d52; @referencing_d52 ||= "Resistive heating"; end
  attr_accessor :referencing_g52 # Default: 20.487336129888465
  def referencing_d53; @referencing_d53 ||= "Oil-fired boiler"; end
  attr_accessor :referencing_g53 # Default: 8.194934451955387
  def referencing_d54; @referencing_d54 ||= "Solid-fuel boiler"; end
  attr_accessor :referencing_g54 # Default: 8.194934451955387
  def referencing_d55; @referencing_d55 ||= "Stirling engine micro-CHP"; end
  attr_accessor :referencing_g55 # Default: 0.0
  def referencing_d56; @referencing_d56 ||= "Fuel-cell micro-CHP"; end
  attr_accessor :referencing_g56 # Default: 0.0
  def referencing_d57; @referencing_d57 ||= "Air-source heat pump"; end
  attr_accessor :referencing_g57 # Default: 0.0
  def referencing_d58; @referencing_d58 ||= "Ground-source heat pump"; end
  attr_accessor :referencing_g58 # Default: 0.0
  def referencing_d59; @referencing_d59 ||= "Geothermal electricity"; end
  attr_accessor :referencing_g59 # Default: 0.0
  def referencing_d60; @referencing_d60 ||= "Community scale gas CHP with local district heating"; end
  attr_accessor :referencing_g60 # Default: 0.0
  def referencing_d61; @referencing_d61 ||= "Community scale solid-fuel CHP with local district heating"; end
  attr_accessor :referencing_g61 # Default: 0.0
  def referencing_d62; @referencing_d62 ||= "Long distance district heating from large power stations"; end
  attr_accessor :referencing_g62 # Default: 0.0
  attr_accessor :referencing_d64 # Default: "H.01"
  def referencing_e64; @referencing_e64 ||= "Heating & cooling"; end
  def referencing_h64; @referencing_h64 ||= sum([[0,0,0,0,multiply(multiply(divide(referencing_g50,referencing_m35),excel_equal?(referencing_d64,referencing_j34)),referencing_j35),0,0,multiply(multiply(divide(referencing_g50,referencing_m35),excel_equal?(referencing_d64,referencing_m34)),referencing_m35),multiply(multiply(divide(referencing_g50,referencing_m35),excel_equal?(referencing_d64,referencing_n34)),referencing_n35)],[0,0,0,0,multiply(multiply(divide(referencing_g51,referencing_m36),excel_equal?(referencing_d64,referencing_j34)),referencing_j36),0,0,multiply(multiply(divide(referencing_g51,referencing_m36),excel_equal?(referencing_d64,referencing_m34)),referencing_m36),multiply(multiply(divide(referencing_g51,referencing_m36),excel_equal?(referencing_d64,referencing_n34)),referencing_n36)],[multiply(multiply(divide(referencing_g52,referencing_m37),excel_equal?(referencing_d64,referencing_f34)),referencing_f37),0,0,0,0,0,0,multiply(multiply(divide(referencing_g52,referencing_m37),excel_equal?(referencing_d64,referencing_m34)),referencing_m37),multiply(multiply(divide(referencing_g52,referencing_m37),excel_equal?(referencing_d64,referencing_n34)),referencing_n37)],[0,0,0,multiply(multiply(divide(referencing_g53,referencing_m38),excel_equal?(referencing_d64,referencing_i34)),referencing_i38),0,0,0,multiply(multiply(divide(referencing_g53,referencing_m38),excel_equal?(referencing_d64,referencing_m34)),referencing_m38),multiply(multiply(divide(referencing_g53,referencing_m38),excel_equal?(referencing_d64,referencing_n34)),referencing_n38)],[0,0,multiply(multiply(divide(referencing_g54,referencing_m39),excel_equal?(referencing_d64,referencing_h34)),referencing_h39),0,0,0,0,multiply(multiply(divide(referencing_g54,referencing_m39),excel_equal?(referencing_d64,referencing_m34)),referencing_m39),multiply(multiply(divide(referencing_g54,referencing_m39),excel_equal?(referencing_d64,referencing_n34)),referencing_n39)],[0,multiply(multiply(divide(referencing_g55,referencing_m40),excel_equal?(referencing_d64,referencing_g34)),referencing_g40),0,0,multiply(multiply(divide(referencing_g55,referencing_m40),excel_equal?(referencing_d64,referencing_j34)),referencing_j40),0,0,multiply(multiply(divide(referencing_g55,referencing_m40),excel_equal?(referencing_d64,referencing_m34)),referencing_m40),multiply(multiply(divide(referencing_g55,referencing_m40),excel_equal?(referencing_d64,referencing_n34)),referencing_n40)],[0,multiply(multiply(divide(referencing_g56,referencing_m41),excel_equal?(referencing_d64,referencing_g34)),referencing_g41),0,0,multiply(multiply(divide(referencing_g56,referencing_m41),excel_equal?(referencing_d64,referencing_j34)),referencing_j41),0,0,multiply(multiply(divide(referencing_g56,referencing_m41),excel_equal?(referencing_d64,referencing_m34)),referencing_m41),multiply(multiply(divide(referencing_g56,referencing_m41),excel_equal?(referencing_d64,referencing_n34)),referencing_n41)],[multiply(multiply(divide(referencing_g57,referencing_m42),excel_equal?(referencing_d64,referencing_f34)),referencing_f42),0,0,0,0,0,multiply(multiply(divide(referencing_g57,referencing_m42),excel_equal?(referencing_d64,referencing_l34)),referencing_l42),multiply(multiply(divide(referencing_g57,referencing_m42),excel_equal?(referencing_d64,referencing_m34)),referencing_m42),0],[multiply(multiply(divide(referencing_g58,referencing_m43),excel_equal?(referencing_d64,referencing_f34)),referencing_f43),0,0,0,0,0,multiply(multiply(divide(referencing_g58,referencing_m43),excel_equal?(referencing_d64,referencing_l34)),referencing_l43),multiply(multiply(divide(referencing_g58,referencing_m43),excel_equal?(referencing_d64,referencing_m34)),referencing_m43),0],[0,0,0,0,0,0,multiply(multiply(divide(referencing_g59,referencing_m44),excel_equal?(referencing_d64,referencing_l34)),referencing_l44),multiply(multiply(divide(referencing_g59,referencing_m44),excel_equal?(referencing_d64,referencing_m34)),referencing_m44),multiply(multiply(divide(referencing_g59,referencing_m44),excel_equal?(referencing_d64,referencing_n34)),referencing_n44)],[0,multiply(multiply(divide(referencing_g60,referencing_m45),excel_equal?(referencing_d64,referencing_g34)),referencing_g45),0,0,multiply(multiply(divide(referencing_g60,referencing_m45),excel_equal?(referencing_d64,referencing_j34)),referencing_j45),0,0,multiply(multiply(divide(referencing_g60,referencing_m45),excel_equal?(referencing_d64,referencing_m34)),referencing_m45),multiply(multiply(divide(referencing_g60,referencing_m45),excel_equal?(referencing_d64,referencing_n34)),referencing_n45)],[0,multiply(multiply(divide(referencing_g61,referencing_m46),excel_equal?(referencing_d64,referencing_g34)),referencing_g46),multiply(multiply(divide(referencing_g61,referencing_m46),excel_equal?(referencing_d64,referencing_h34)),referencing_h46),0,0,0,0,multiply(multiply(divide(referencing_g61,referencing_m46),excel_equal?(referencing_d64,referencing_m34)),referencing_m46),multiply(multiply(divide(referencing_g61,referencing_m46),excel_equal?(referencing_d64,referencing_n34)),referencing_n46)],[0,0,0,0,0,multiply(multiply(divide(referencing_g62,referencing_m47),excel_equal?(referencing_d64,referencing_k34)),referencing_k47),0,multiply(multiply(divide(referencing_g62,referencing_m47),excel_equal?(referencing_d64,referencing_m34)),referencing_m47),multiply(multiply(divide(referencing_g62,referencing_m47),excel_equal?(referencing_d64,referencing_n34)),referencing_n47)]]); end
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
  def tables_b5; @tables_b5 ||= common7; end
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
  def tables_c11; @tables_c11 ||= common7; end
  def tables_e11; @tables_e11 ||= tables_b5; end
  def tables_f11; @tables_f11 ||= tables_c5; end
  def tables_g11; @tables_g11 ||= nil; end
  def tables_c12; @tables_c12 ||= tables_b5; end
  def tables_c13; @tables_c13 ||= common9; end
  def tables_c14; @tables_c14 ||= common9; end
  def s_innapropriate_sheet_name__c4; @s_innapropriate_sheet_name__c4 ||= valuetypes_a3; end
  def common0; @common0 ||= add(referencing_a4,1.0); end
  def common1; @common1 ||= sum([[tables_b5,tables_c5,nil]]); end
  def common3; @common3 ||= index([[referencing_c16,referencing_d16,referencing_e16,referencing_f16],[referencing_c17,referencing_d17,referencing_e17,referencing_f17],[referencing_c18,referencing_d18,referencing_e18,referencing_f18],[referencing_c19,referencing_d19,referencing_e19,referencing_f19]],nil,excel_match(referencing_c22,[[referencing_c15,referencing_d15,referencing_e15,referencing_f15]],0.0)); end
  def common7; @common7 ||= sum([[tables_b3],[tables_b4]]); end
  def common9; @common9 ||= sum([[tables_b3,tables_c3,tables_d3],[tables_b4,tables_c4,tables_d4]]); end


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
    @referencing_c15 = 1.0
    @referencing_d15 = 2.0
    @referencing_e15 = 3.0
    @referencing_f15 = 4.0
    @referencing_c16 = 1.4535833325868115
    @referencing_d16 = 1.4535833325868115
    @referencing_e16 = 1.511726665890284
    @referencing_f16 = 1.5407983325420203
    @referencing_c17 = 9.054545454545455
    @referencing_d17 = 12.0
    @referencing_e17 = 18.0
    @referencing_f17 = 18.0
    @referencing_c18 = 0.3681150635671386
    @referencing_d18 = 0.3681150635671386
    @referencing_e18 = 0.40588480110308967
    @referencing_f18 = 0.42190146532760275
    @referencing_c19 = 0.651
    @referencing_d19 = 0.651
    @referencing_e19 = 0.651
    @referencing_f19 = 0.651
    @referencing_c22 = 4.0
    @referencing_f34 = "V.01"
    @referencing_g34 = "V.02"
    @referencing_h34 = "V.03"
    @referencing_i34 = "V.04"
    @referencing_j34 = "V.05"
    @referencing_k34 = "V.07"
    @referencing_l34 = "R.07"
    @referencing_m34 = "H.01"
    @referencing_n34 = "X.01"
    @referencing_j35 = -1.0
    @referencing_m35 = 0.76
    @referencing_n35 = 0.24
    @referencing_j36 = -1.0
    @referencing_m36 = 0.91
    @referencing_n36 = 0.09
    @referencing_f37 = -1.0
    @referencing_m37 = 1.0
    @referencing_n37 = 0.0
    @referencing_i38 = -1.0
    @referencing_m38 = 0.97
    @referencing_n38 = 0.03
    @referencing_h39 = -1.0
    @referencing_m39 = 0.87
    @referencing_n39 = 0.13
    @referencing_g40 = 0.225
    @referencing_j40 = -1.0
    @referencing_m40 = 0.63
    @referencing_n40 = 0.145
    @referencing_g41 = 0.45
    @referencing_j41 = -1.0
    @referencing_m41 = 0.45
    @referencing_n41 = 0.1
    @referencing_f42 = -1.0
    @referencing_l42 = -1.0
    @referencing_m42 = 2.0
    @referencing_f43 = -1.0
    @referencing_l43 = -2.0
    @referencing_m43 = 3.0
    @referencing_l44 = -1.0
    @referencing_m44 = 0.85
    @referencing_n44 = 0.15
    @referencing_g45 = 0.38
    @referencing_j45 = -1.0
    @referencing_m45 = 0.38
    @referencing_n45 = 0.24
    @referencing_g46 = 0.17
    @referencing_h46 = -1.0
    @referencing_m46 = 0.57
    @referencing_n46 = 0.26
    @referencing_k47 = -1.0
    @referencing_m47 = 0.9
    @referencing_n47 = 0.1
    @referencing_g50 = 137.26515207025273
    @referencing_g51 = 30.731004194832696
    @referencing_g52 = 20.487336129888465
    @referencing_g53 = 8.194934451955387
    @referencing_g54 = 8.194934451955387
    @referencing_g55 = 0.0
    @referencing_g56 = 0.0
    @referencing_g57 = 0.0
    @referencing_g58 = 0.0
    @referencing_g59 = 0.0
    @referencing_g60 = 0.0
    @referencing_g61 = 0.0
    @referencing_g62 = 0.0
    @referencing_d64 = "H.01"
    @tables_b2 = "ColA"
    @tables_c2 = "ColB"
    @tables_d2 = "Column1"
    @tables_b3 = 1.0
    @tables_c3 = "A"
    @tables_b4 = 2.0
    @tables_c4 = "B"
  end

end
