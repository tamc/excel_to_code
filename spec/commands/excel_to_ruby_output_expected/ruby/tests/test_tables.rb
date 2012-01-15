# Test for Tables
require 'test/unit'
require_relative '../examplespreadsheet'

module ExampleSpreadsheet
class TestTables < Test::Unit::TestCase
  def worksheet; Tables.new; end
  def test_b2; assert_equal("ColA",worksheet.b2); end
  def test_c2; assert_equal("ColB",worksheet.c2); end
  def test_d2; assert_equal("Column1",worksheet.d2); end
  def test_g2; assert_equal(:value,worksheet.g2); end
  def test_b3; assert_equal(1,worksheet.b3); end
  def test_c3; assert_equal("A",worksheet.c3); end
  def test_d3; assert_equal(:value,worksheet.d3); end
  def test_b4; assert_equal(2,worksheet.b4); end
  def test_c4; assert_equal("B",worksheet.c4); end
  def test_d4; assert_equal(:value,worksheet.d4); end
  def test_f4; assert_equal("B",worksheet.f4); end
  def test_g4; assert_equal(:value,worksheet.g4); end
  def test_h4; assert_equal(:value,worksheet.h4); end
  def test_b5; assert_equal(3,worksheet.b5); end
  def test_c5; assert_equal(0,worksheet.c5); end
  def test_e6; assert_equal(:value,worksheet.e6); end
  def test_e7; assert_equal(:value,worksheet.e7); end
  def test_e8; assert_equal(:value,worksheet.e8); end
  def test_e9; assert_equal(:value,worksheet.e9); end
  def test_c10; assert_equal(0,worksheet.c10); end
  def test_e10; assert_equal(:value,worksheet.e10); end
  def test_c11; assert_equal(3,worksheet.c11); end
  def test_c12; assert_equal(3,worksheet.c12); end
end
end
