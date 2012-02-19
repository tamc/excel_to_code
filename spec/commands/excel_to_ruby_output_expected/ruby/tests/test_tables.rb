# coding: utf-8
# Test for Tables
require 'test/unit'
require_relative '../examplespreadsheet'

module ExampleSpreadsheet
class TestTables < Test::Unit::TestCase
  def spreadsheet; $spreadsheet ||= Spreadsheet.new; end
  def worksheet; @worksheet ||= spreadsheet.tables; end
  def test_a1; assert_in_epsilon(12,worksheet.a1); end
  def test_b2; assert_equal("ColA",worksheet.b2); end
  def test_c2; assert_equal("ColB",worksheet.c2); end
  def test_d2; assert_equal("Column1",worksheet.d2); end
  def test_b3; assert_in_epsilon(1,worksheet.b3); end
  def test_c3; assert_equal("A",worksheet.c3); end
  def test_d3; assert_equal("1A",worksheet.d3); end
  def test_b4; assert_in_epsilon(2,worksheet.b4); end
  def test_c4; assert_equal("B",worksheet.c4); end
  def test_d4; assert_equal("2B",worksheet.d4); end
  def test_f4; assert_equal("B",worksheet.f4); end
  def test_g4; assert_in_epsilon(3,worksheet.g4); end
  def test_h4; assert_in_epsilon(1,worksheet.h4); end
  def test_b5; assert_in_epsilon(3,worksheet.b5); end
  def test_c5; assert_in_epsilon(0,worksheet.c5 || 0); end
  def test_e6; assert_equal("ColA",worksheet.e6); end
  def test_f6; assert_equal("ColB",worksheet.f6); end
  def test_g6; assert_equal("Column1",worksheet.g6); end
  def test_e7; assert_in_epsilon(3,worksheet.e7); end
  def test_f7; assert_in_epsilon(0,worksheet.f7 || 0); end
  def test_g7; assert_in_epsilon(0,worksheet.g7 || 0); end
  def test_e8; assert_equal("ColA",worksheet.e8); end
  def test_f8; assert_equal("ColB",worksheet.f8); end
  def test_g8; assert_equal("Column1",worksheet.g8); end
  def test_e9; assert_in_epsilon(1,worksheet.e9); end
  def test_f9; assert_equal("A",worksheet.f9); end
  def test_g9; assert_equal("1A",worksheet.g9); end
  def test_c10; assert_in_epsilon(3,worksheet.c10); end
  def test_e10; assert_in_epsilon(2,worksheet.e10); end
  def test_f10; assert_equal("B",worksheet.f10); end
  def test_g10; assert_equal("2B",worksheet.g10); end
  def test_c11; assert_in_epsilon(3,worksheet.c11); end
  def test_e11; assert_in_epsilon(3,worksheet.e11); end
  def test_f11; assert_in_epsilon(0,worksheet.f11 || 0); end
  def test_g11; assert_in_epsilon(0,worksheet.g11 || 0); end
  def test_c12; assert_in_epsilon(3,worksheet.c12); end
  def test_c13; assert_in_epsilon(3,worksheet.c13); end
  def test_c14; assert_in_epsilon(3,worksheet.c14); end
end
end
