# coding: utf-8
# All tests for /Users/tamc/Documents/github/excel_to_code/spec/test_data/ExampleSpreadsheet.xlsx
require 'test/unit'
require_relative 'examplespreadsheet'

class TestExampleSpreadsheet < Test::Unit::TestCase
  def worksheet; @worksheet ||= ExampleSpreadsheet.new; end
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
  def test_tables_a1; assert_in_delta(0.0, worksheet.tables_a1, 0.002); end
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
  def test_tables_c5; assert_in_delta(0.0, worksheet.tables_c5, 0.002); end
  def test_tables_e6; assert_equal("ColA", worksheet.tables_e6); end
  def test_tables_f6; assert_equal("ColB", worksheet.tables_f6); end
  def test_tables_g6; assert_equal("Column1", worksheet.tables_g6); end
  def test_tables_e7; assert_in_epsilon(3.0, worksheet.tables_e7, 0.002); end
  def test_tables_f7; assert_in_delta(0.0, worksheet.tables_f7, 0.002); end
  def test_tables_g7; assert_in_delta(0.0, worksheet.tables_g7, 0.002); end
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
  def test_tables_f11; assert_in_delta(0.0, worksheet.tables_f11, 0.002); end
  def test_tables_g11; assert_in_delta(0.0, worksheet.tables_g11, 0.002); end
  def test_tables_c12; assert_in_epsilon(3.0, worksheet.tables_c12, 0.002); end
  def test_tables_c13; assert_in_epsilon(3.0, worksheet.tables_c13, 0.002); end
  def test_tables_c14; assert_in_epsilon(3.0, worksheet.tables_c14, 0.002); end
  def test_s_innapropriate_sheet_name__c4; assert_in_delta(1.0, worksheet.s_innapropriate_sheet_name__c4, 0.002); end
end
