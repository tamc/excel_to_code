# coding: utf-8
# All tests for /Users/tamc/Documents/github/excel_to_code/spec/test_data/ExampleSpreadsheet.xlsx
require 'test/unit'
require_relative 'rubyexamplespreadsheet'

class TestRubyExampleSpreadsheet < Test::Unit::TestCase
  def worksheet; @worksheet ||= RubyExampleSpreadsheet.new; end
  def test_valuetypes_a1; assert_equal(true, worksheet.valuetypes_a1); end
  def test_valuetypes_a2; assert_equal("Hello", worksheet.valuetypes_a2); end
  def test_valuetypes_a3; assert_in_delta(1.0, worksheet.valuetypes_a3, 0.002); end
  def test_valuetypes_a4; assert_in_epsilon(3.1415, worksheet.valuetypes_a4, 0.002); end
  def test_valuetypes_a5; assert_equal(:name, worksheet.valuetypes_a5); end
  def test_valuetypes_a6; assert_equal("Hello", worksheet.valuetypes_a6); end
  def test_formulaetypes_a1; assert_equal("Simple", worksheet.formulaetypes_a1); end
  def test_formulaetypes_b1; assert_in_epsilon(2.0, worksheet.formulaetypes_b1, 0.002); end
  def test_formulaetypes_a2; assert_equal("Sharing", worksheet.formulaetypes_a2); end
  def test_formulaetypes_b2; assert_in_epsilon(267.7467614837482, worksheet.formulaetypes_b2, 0.002); end
  def test_formulaetypes_a3; assert_equal("Shared", worksheet.formulaetypes_a3); end
  def test_formulaetypes_b3; assert_in_epsilon(267.7467614837482, worksheet.formulaetypes_b3, 0.002); end
  def test_formulaetypes_a4; assert_equal("Shared", worksheet.formulaetypes_a4); end
  def test_formulaetypes_b4; assert_in_epsilon(267.7467614837482, worksheet.formulaetypes_b4, 0.002); end
  def test_formulaetypes_a5; assert_equal("Array (single)", worksheet.formulaetypes_a5); end
  def test_formulaetypes_b5; assert_in_epsilon(2.0, worksheet.formulaetypes_b5, 0.002); end
  def test_formulaetypes_a6; assert_equal("Arraying (multiple)", worksheet.formulaetypes_a6); end
  def test_formulaetypes_b6; assert_equal("Not Eight", worksheet.formulaetypes_b6); end
  def test_formulaetypes_a7; assert_equal("Arrayed (multiple)", worksheet.formulaetypes_a7); end
  def test_formulaetypes_b7; assert_equal("Not Eight", worksheet.formulaetypes_b7); end
  def test_formulaetypes_a8; assert_equal("Arrayed (multiple)", worksheet.formulaetypes_a8); end
  def test_formulaetypes_b8; assert_equal("Not Eight", worksheet.formulaetypes_b8); end
  def test_ranges_b1; assert_equal("This sheet", worksheet.ranges_b1); end
  def test_ranges_c1; assert_equal("Other sheet", worksheet.ranges_c1); end
  def test_ranges_a2; assert_equal("Standard", worksheet.ranges_a2); end
  def test_ranges_b2; assert_in_epsilon(6.0, worksheet.ranges_b2, 0.002); end
  def test_ranges_c2; assert_in_epsilon(4.141500000000001, worksheet.ranges_c2, 0.002); end
  def test_ranges_a3; assert_equal("Column", worksheet.ranges_a3); end
  def test_ranges_b3; assert_in_epsilon(6.0, worksheet.ranges_b3, 0.002); end
  def test_ranges_c3; assert_equal(:name, worksheet.ranges_c3); end
  def test_ranges_a4; assert_equal("Row", worksheet.ranges_a4); end
  def test_ranges_b4; assert_in_epsilon(6.0, worksheet.ranges_b4, 0.002); end
  def test_ranges_c4; assert_in_epsilon(3.1415, worksheet.ranges_c4, 0.002); end
  def test_ranges_f4; assert_in_delta(1.0, worksheet.ranges_f4, 0.002); end
  def test_ranges_e5; assert_in_delta(1.0, worksheet.ranges_e5, 0.002); end
  def test_ranges_f5; assert_in_epsilon(2.0, worksheet.ranges_f5, 0.002); end
  def test_ranges_g5; assert_in_epsilon(3.0, worksheet.ranges_g5, 0.002); end
  def test_ranges_f6; assert_in_epsilon(3.0, worksheet.ranges_f6, 0.002); end
  def test_referencing_a1; assert_in_epsilon(12.0, worksheet.referencing_a1, 0.002); end
  def test_referencing_a2; assert_in_epsilon(12.0, worksheet.referencing_a2, 0.002); end
  def test_referencing_a4; assert_in_epsilon(10.0, worksheet.referencing_a4, 0.002); end
  def test_referencing_b4; assert_in_epsilon(11.0, worksheet.referencing_b4, 0.002); end
  def test_referencing_c4; assert_in_epsilon(12.0, worksheet.referencing_c4, 0.002); end
  def test_referencing_a5; assert_in_epsilon(3.0, worksheet.referencing_a5, 0.002); end
  def test_referencing_b8; assert_in_epsilon(12.0, worksheet.referencing_b8, 0.002); end
  def test_referencing_b9; assert_in_epsilon(3.0, worksheet.referencing_b9, 0.002); end
  def test_referencing_b11; assert_equal("Named", worksheet.referencing_b11); end
  def test_referencing_c11; assert_equal("Reference", worksheet.referencing_c11); end
  def test_referencing_c15; assert_in_delta(1.0, worksheet.referencing_c15, 0.002); end
  def test_referencing_d15; assert_in_epsilon(2.0, worksheet.referencing_d15, 0.002); end
  def test_referencing_e15; assert_in_epsilon(3.0, worksheet.referencing_e15, 0.002); end
  def test_referencing_f15; assert_in_epsilon(4.0, worksheet.referencing_f15, 0.002); end
  def test_referencing_c16; assert_in_epsilon(1.4535833325868115, worksheet.referencing_c16, 0.002); end
  def test_referencing_d16; assert_in_epsilon(1.4535833325868115, worksheet.referencing_d16, 0.002); end
  def test_referencing_e16; assert_in_epsilon(1.511726665890284, worksheet.referencing_e16, 0.002); end
  def test_referencing_f16; assert_in_epsilon(1.5407983325420203, worksheet.referencing_f16, 0.002); end
  def test_referencing_c17; assert_in_epsilon(9.054545454545455, worksheet.referencing_c17, 0.002); end
  def test_referencing_d17; assert_in_epsilon(12.0, worksheet.referencing_d17, 0.002); end
  def test_referencing_e17; assert_in_epsilon(18.0, worksheet.referencing_e17, 0.002); end
  def test_referencing_f17; assert_in_epsilon(18.0, worksheet.referencing_f17, 0.002); end
  def test_referencing_c18; assert_in_delta(0.3681150635671386, worksheet.referencing_c18, 0.002); end
  def test_referencing_d18; assert_in_delta(0.3681150635671386, worksheet.referencing_d18, 0.002); end
  def test_referencing_e18; assert_in_delta(0.40588480110308967, worksheet.referencing_e18, 0.002); end
  def test_referencing_f18; assert_in_delta(0.42190146532760275, worksheet.referencing_f18, 0.002); end
  def test_referencing_c19; assert_in_delta(0.651, worksheet.referencing_c19, 0.002); end
  def test_referencing_d19; assert_in_delta(0.651, worksheet.referencing_d19, 0.002); end
  def test_referencing_e19; assert_in_delta(0.651, worksheet.referencing_e19, 0.002); end
  def test_referencing_f19; assert_in_delta(0.651, worksheet.referencing_f19, 0.002); end
  def test_referencing_c22; assert_in_epsilon(4.0, worksheet.referencing_c22, 0.002); end
  def test_referencing_d22; assert_in_epsilon(1.5407983325420203, worksheet.referencing_d22, 0.002); end
  def test_referencing_d23; assert_in_epsilon(18.0, worksheet.referencing_d23, 0.002); end
  def test_referencing_d24; assert_in_delta(0.42190146532760275, worksheet.referencing_d24, 0.002); end
  def test_referencing_d25; assert_in_delta(0.651, worksheet.referencing_d25, 0.002); end
  def test_referencing_c31; assert_equal("Technology efficiencies -- hot water -- annual mean", worksheet.referencing_c31); end
  def test_referencing_o31; assert_equal("% of input energy", worksheet.referencing_o31); end
  def test_referencing_f33; assert_equal("Electricity (delivered to end user)", worksheet.referencing_f33); end
  def test_referencing_g33; assert_equal("Electricity (supplied to grid)", worksheet.referencing_g33); end
  def test_referencing_h33; assert_equal("Solid hydrocarbons", worksheet.referencing_h33); end
  def test_referencing_i33; assert_equal("Liquid hydrocarbons", worksheet.referencing_i33); end
  def test_referencing_j33; assert_equal("Gaseous hydrocarbons", worksheet.referencing_j33); end
  def test_referencing_k33; assert_equal("Heat transport", worksheet.referencing_k33); end
  def test_referencing_l33; assert_equal("Environmental heat", worksheet.referencing_l33); end
  def test_referencing_m33; assert_equal("Heating & cooling", worksheet.referencing_m33); end
  def test_referencing_n33; assert_equal("Conversion losses", worksheet.referencing_n33); end
  def test_referencing_o33; assert_equal("Balance", worksheet.referencing_o33); end
  def test_referencing_c34; assert_equal("Code", worksheet.referencing_c34); end
  def test_referencing_d34; assert_equal("Technology", worksheet.referencing_d34); end
  def test_referencing_e34; assert_equal("Notes", worksheet.referencing_e34); end
  def test_referencing_f34; assert_equal("V.01", worksheet.referencing_f34); end
  def test_referencing_g34; assert_equal("V.02", worksheet.referencing_g34); end
  def test_referencing_h34; assert_equal("V.03", worksheet.referencing_h34); end
  def test_referencing_i34; assert_equal("V.04", worksheet.referencing_i34); end
  def test_referencing_j34; assert_equal("V.05", worksheet.referencing_j34); end
  def test_referencing_k34; assert_equal("V.07", worksheet.referencing_k34); end
  def test_referencing_l34; assert_equal("R.07", worksheet.referencing_l34); end
  def test_referencing_m34; assert_equal("H.01", worksheet.referencing_m34); end
  def test_referencing_n34; assert_equal("X.01", worksheet.referencing_n34); end
  def test_referencing_c35; assert_in_delta(1.0, worksheet.referencing_c35, 0.002); end
  def test_referencing_d35; assert_equal("Gas boiler (old)", worksheet.referencing_d35); end
  def test_referencing_j35; assert_in_delta(-1.0, worksheet.referencing_j35, 0.002); end
  def test_referencing_m35; assert_in_delta(0.76, worksheet.referencing_m35, 0.002); end
  def test_referencing_n35; assert_in_delta(0.24, worksheet.referencing_n35, 0.002); end
  def test_referencing_o35; assert_in_delta(0.0, (worksheet.referencing_o35||0), 0.002); end
  def test_referencing_c36; assert_in_epsilon(2.0, worksheet.referencing_c36, 0.002); end
  def test_referencing_d36; assert_equal("Gas boiler (new)", worksheet.referencing_d36); end
  def test_referencing_j36; assert_in_delta(-1.0, worksheet.referencing_j36, 0.002); end
  def test_referencing_m36; assert_in_delta(0.91, worksheet.referencing_m36, 0.002); end
  def test_referencing_n36; assert_in_delta(0.09, worksheet.referencing_n36, 0.002); end
  def test_referencing_o36; assert_in_delta(0.0, (worksheet.referencing_o36||0), 0.002); end
  def test_referencing_c37; assert_in_epsilon(3.0, worksheet.referencing_c37, 0.002); end
  def test_referencing_d37; assert_equal("Resistive heating", worksheet.referencing_d37); end
  def test_referencing_f37; assert_in_delta(-1.0, worksheet.referencing_f37, 0.002); end
  def test_referencing_m37; assert_in_delta(1.0, worksheet.referencing_m37, 0.002); end
  def test_referencing_n37; assert_in_delta(0.0, (worksheet.referencing_n37||0), 0.002); end
  def test_referencing_o37; assert_in_delta(0.0, (worksheet.referencing_o37||0), 0.002); end
  def test_referencing_c38; assert_in_epsilon(4.0, worksheet.referencing_c38, 0.002); end
  def test_referencing_d38; assert_equal("Oil-fired boiler", worksheet.referencing_d38); end
  def test_referencing_i38; assert_in_delta(-1.0, worksheet.referencing_i38, 0.002); end
  def test_referencing_m38; assert_in_delta(0.97, worksheet.referencing_m38, 0.002); end
  def test_referencing_n38; assert_in_delta(0.03, worksheet.referencing_n38, 0.002); end
  def test_referencing_o38; assert_in_delta(-2.7755575615628914e-17, worksheet.referencing_o38, 0.002); end
  def test_referencing_c39; assert_in_epsilon(5.0, worksheet.referencing_c39, 0.002); end
  def test_referencing_d39; assert_equal("Solid-fuel boiler", worksheet.referencing_d39); end
  def test_referencing_e39; assert_equal("[2]", worksheet.referencing_e39); end
  def test_referencing_h39; assert_in_delta(-1.0, worksheet.referencing_h39, 0.002); end
  def test_referencing_m39; assert_in_delta(0.87, worksheet.referencing_m39, 0.002); end
  def test_referencing_n39; assert_in_delta(0.13, worksheet.referencing_n39, 0.002); end
  def test_referencing_o39; assert_in_delta(0.0, (worksheet.referencing_o39||0), 0.002); end
  def test_referencing_c40; assert_in_epsilon(6.0, worksheet.referencing_c40, 0.002); end
  def test_referencing_d40; assert_equal("Stirling engine micro-CHP", worksheet.referencing_d40); end
  def test_referencing_e40; assert_equal("[3]", worksheet.referencing_e40); end
  def test_referencing_g40; assert_in_delta(0.225, worksheet.referencing_g40, 0.002); end
  def test_referencing_j40; assert_in_delta(-1.0, worksheet.referencing_j40, 0.002); end
  def test_referencing_m40; assert_in_delta(0.63, worksheet.referencing_m40, 0.002); end
  def test_referencing_n40; assert_in_delta(0.145, worksheet.referencing_n40, 0.002); end
  def test_referencing_o40; assert_in_delta(0.0, (worksheet.referencing_o40||0), 0.002); end
  def test_referencing_c41; assert_in_epsilon(7.0, worksheet.referencing_c41, 0.002); end
  def test_referencing_d41; assert_equal("Fuel-cell micro-CHP", worksheet.referencing_d41); end
  def test_referencing_e41; assert_equal("[3]", worksheet.referencing_e41); end
  def test_referencing_g41; assert_in_delta(0.45, worksheet.referencing_g41, 0.002); end
  def test_referencing_j41; assert_in_delta(-1.0, worksheet.referencing_j41, 0.002); end
  def test_referencing_m41; assert_in_delta(0.45, worksheet.referencing_m41, 0.002); end
  def test_referencing_n41; assert_in_delta(0.1, worksheet.referencing_n41, 0.002); end
  def test_referencing_o41; assert_in_delta(0.0, (worksheet.referencing_o41||0), 0.002); end
  def test_referencing_c42; assert_in_epsilon(8.0, worksheet.referencing_c42, 0.002); end
  def test_referencing_d42; assert_equal("Air-source heat pump", worksheet.referencing_d42); end
  def test_referencing_f42; assert_in_delta(-1.0, worksheet.referencing_f42, 0.002); end
  def test_referencing_l42; assert_in_delta(-1.0, worksheet.referencing_l42, 0.002); end
  def test_referencing_m42; assert_in_epsilon(2.0, worksheet.referencing_m42, 0.002); end
  def test_referencing_o42; assert_in_delta(0.0, (worksheet.referencing_o42||0), 0.002); end
  def test_referencing_c43; assert_in_epsilon(9.0, worksheet.referencing_c43, 0.002); end
  def test_referencing_d43; assert_equal("Ground-source heat pump", worksheet.referencing_d43); end
  def test_referencing_f43; assert_in_delta(-1.0, worksheet.referencing_f43, 0.002); end
  def test_referencing_l43; assert_in_epsilon(-2.0, worksheet.referencing_l43, 0.002); end
  def test_referencing_m43; assert_in_epsilon(3.0, worksheet.referencing_m43, 0.002); end
  def test_referencing_o43; assert_in_delta(0.0, (worksheet.referencing_o43||0), 0.002); end
  def test_referencing_c44; assert_in_epsilon(10.0, worksheet.referencing_c44, 0.002); end
  def test_referencing_d44; assert_equal("Geothermal electricity", worksheet.referencing_d44); end
  def test_referencing_l44; assert_in_delta(-1.0, worksheet.referencing_l44, 0.002); end
  def test_referencing_m44; assert_in_delta(0.85, worksheet.referencing_m44, 0.002); end
  def test_referencing_n44; assert_in_delta(0.15, worksheet.referencing_n44, 0.002); end
  def test_referencing_o44; assert_in_delta(0.0, (worksheet.referencing_o44||0), 0.002); end
  def test_referencing_c45; assert_in_epsilon(11.0, worksheet.referencing_c45, 0.002); end
  def test_referencing_d45; assert_equal("Community scale gas CHP with local district heating", worksheet.referencing_d45); end
  def test_referencing_g45; assert_in_delta(0.38, worksheet.referencing_g45, 0.002); end
  def test_referencing_j45; assert_in_delta(-1.0, worksheet.referencing_j45, 0.002); end
  def test_referencing_m45; assert_in_delta(0.38, worksheet.referencing_m45, 0.002); end
  def test_referencing_n45; assert_in_delta(0.24, worksheet.referencing_n45, 0.002); end
  def test_referencing_o45; assert_in_delta(0.0, (worksheet.referencing_o45||0), 0.002); end
  def test_referencing_c46; assert_in_epsilon(12.0, worksheet.referencing_c46, 0.002); end
  def test_referencing_d46; assert_equal("Community scale solid-fuel CHP with local district heating", worksheet.referencing_d46); end
  def test_referencing_g46; assert_in_delta(0.17, worksheet.referencing_g46, 0.002); end
  def test_referencing_h46; assert_in_delta(-1.0, worksheet.referencing_h46, 0.002); end
  def test_referencing_m46; assert_in_delta(0.57, worksheet.referencing_m46, 0.002); end
  def test_referencing_n46; assert_in_delta(0.26, worksheet.referencing_n46, 0.002); end
  def test_referencing_o46; assert_in_delta(0.0, (worksheet.referencing_o46||0), 0.002); end
  def test_referencing_c47; assert_in_epsilon(13.0, worksheet.referencing_c47, 0.002); end
  def test_referencing_d47; assert_equal("Long distance district heating from large power stations", worksheet.referencing_d47); end
  def test_referencing_e47; assert_equal("[6]", worksheet.referencing_e47); end
  def test_referencing_k47; assert_in_delta(-1.0, worksheet.referencing_k47, 0.002); end
  def test_referencing_m47; assert_in_delta(0.9, worksheet.referencing_m47, 0.002); end
  def test_referencing_n47; assert_in_delta(0.1, worksheet.referencing_n47, 0.002); end
  def test_referencing_o47; assert_in_delta(0.0, (worksheet.referencing_o47||0), 0.002); end
  def test_referencing_d50; assert_equal("Gas boiler (old)", worksheet.referencing_d50); end
  def test_referencing_g50; assert_in_epsilon(137.26515207025273, worksheet.referencing_g50, 0.002); end
  def test_referencing_d51; assert_equal("Gas boiler (new)", worksheet.referencing_d51); end
  def test_referencing_g51; assert_in_epsilon(30.731004194832696, worksheet.referencing_g51, 0.002); end
  def test_referencing_d52; assert_equal("Resistive heating", worksheet.referencing_d52); end
  def test_referencing_g52; assert_in_epsilon(20.487336129888465, worksheet.referencing_g52, 0.002); end
  def test_referencing_d53; assert_equal("Oil-fired boiler", worksheet.referencing_d53); end
  def test_referencing_g53; assert_in_epsilon(8.194934451955387, worksheet.referencing_g53, 0.002); end
  def test_referencing_d54; assert_equal("Solid-fuel boiler", worksheet.referencing_d54); end
  def test_referencing_g54; assert_in_epsilon(8.194934451955387, worksheet.referencing_g54, 0.002); end
  def test_referencing_d55; assert_equal("Stirling engine micro-CHP", worksheet.referencing_d55); end
  def test_referencing_g55; assert_in_delta(0.0, (worksheet.referencing_g55||0), 0.002); end
  def test_referencing_d56; assert_equal("Fuel-cell micro-CHP", worksheet.referencing_d56); end
  def test_referencing_g56; assert_in_delta(0.0, (worksheet.referencing_g56||0), 0.002); end
  def test_referencing_d57; assert_equal("Air-source heat pump", worksheet.referencing_d57); end
  def test_referencing_g57; assert_in_delta(0.0, (worksheet.referencing_g57||0), 0.002); end
  def test_referencing_d58; assert_equal("Ground-source heat pump", worksheet.referencing_d58); end
  def test_referencing_g58; assert_in_delta(0.0, (worksheet.referencing_g58||0), 0.002); end
  def test_referencing_d59; assert_equal("Geothermal electricity", worksheet.referencing_d59); end
  def test_referencing_g59; assert_in_delta(0.0, (worksheet.referencing_g59||0), 0.002); end
  def test_referencing_d60; assert_equal("Community scale gas CHP with local district heating", worksheet.referencing_d60); end
  def test_referencing_g60; assert_in_delta(0.0, (worksheet.referencing_g60||0), 0.002); end
  def test_referencing_d61; assert_equal("Community scale solid-fuel CHP with local district heating", worksheet.referencing_d61); end
  def test_referencing_g61; assert_in_delta(0.0, (worksheet.referencing_g61||0), 0.002); end
  def test_referencing_d62; assert_equal("Long distance district heating from large power stations", worksheet.referencing_d62); end
  def test_referencing_g62; assert_in_delta(0.0, (worksheet.referencing_g62||0), 0.002); end
  def test_referencing_d64; assert_equal("H.01", worksheet.referencing_d64); end
  def test_referencing_e64; assert_equal("Heating & cooling", worksheet.referencing_e64); end
  def test_referencing_h64; assert_in_epsilon(204.87336129888465, worksheet.referencing_h64, 0.002); end
  def test_tables_a1; assert_in_delta(0.0, (worksheet.tables_a1||0), 0.002); end
  def test_tables_b2; assert_equal("ColA", worksheet.tables_b2); end
  def test_tables_c2; assert_equal("ColB", worksheet.tables_c2); end
  def test_tables_d2; assert_equal("Column1", worksheet.tables_d2); end
  def test_tables_b3; assert_in_delta(1.0, worksheet.tables_b3, 0.002); end
  def test_tables_c3; assert_equal("A", worksheet.tables_c3); end
  def test_tables_d3; assert_equal("1A", worksheet.tables_d3); end
  def test_tables_b4; assert_in_epsilon(2.0, worksheet.tables_b4, 0.002); end
  def test_tables_c4; assert_equal("B", worksheet.tables_c4); end
  def test_tables_d4; assert_equal("2B", worksheet.tables_d4); end
  def test_tables_f4; assert_equal("B", worksheet.tables_f4); end
  def test_tables_g4; assert_in_epsilon(3.0, worksheet.tables_g4, 0.002); end
  def test_tables_h4; assert_in_delta(1.0, worksheet.tables_h4, 0.002); end
  def test_tables_b5; assert_in_epsilon(3.0, worksheet.tables_b5, 0.002); end
  def test_tables_c5; assert_in_delta(0.0, (worksheet.tables_c5||0), 0.002); end
  def test_tables_e6; assert_equal("ColA", worksheet.tables_e6); end
  def test_tables_f6; assert_equal("ColB", worksheet.tables_f6); end
  def test_tables_g6; assert_equal("Column1", worksheet.tables_g6); end
  def test_tables_e7; assert_in_epsilon(3.0, worksheet.tables_e7, 0.002); end
  def test_tables_f7; assert_in_delta(0.0, (worksheet.tables_f7||0), 0.002); end
  def test_tables_g7; assert_in_delta(0.0, (worksheet.tables_g7||0), 0.002); end
  def test_tables_e8; assert_equal("ColA", worksheet.tables_e8); end
  def test_tables_f8; assert_equal("ColB", worksheet.tables_f8); end
  def test_tables_g8; assert_equal("Column1", worksheet.tables_g8); end
  def test_tables_e9; assert_in_delta(1.0, worksheet.tables_e9, 0.002); end
  def test_tables_f9; assert_equal("A", worksheet.tables_f9); end
  def test_tables_g9; assert_equal("1A", worksheet.tables_g9); end
  def test_tables_c10; assert_in_epsilon(3.0, worksheet.tables_c10, 0.002); end
  def test_tables_e10; assert_in_epsilon(2.0, worksheet.tables_e10, 0.002); end
  def test_tables_f10; assert_equal("B", worksheet.tables_f10); end
  def test_tables_g10; assert_equal("2B", worksheet.tables_g10); end
  def test_tables_c11; assert_in_epsilon(3.0, worksheet.tables_c11, 0.002); end
  def test_tables_e11; assert_in_epsilon(3.0, worksheet.tables_e11, 0.002); end
  def test_tables_f11; assert_in_delta(0.0, (worksheet.tables_f11||0), 0.002); end
  def test_tables_g11; assert_in_delta(0.0, (worksheet.tables_g11||0), 0.002); end
  def test_tables_c12; assert_in_epsilon(3.0, worksheet.tables_c12, 0.002); end
  def test_tables_c13; assert_in_epsilon(3.0, worksheet.tables_c13, 0.002); end
  def test_tables_c14; assert_in_epsilon(3.0, worksheet.tables_c14, 0.002); end
  def test_s_innapropriate_sheet_name__c4; assert_in_delta(1.0, worksheet.s_innapropriate_sheet_name__c4, 0.002); end
end
